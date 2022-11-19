using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace SendNewsLetter
{
    /// <summary>File utility functions</summary>
    public static class FileUtil
    {
        /// <summary>Returns the path for the executable file that started the application</summary>
        public static string GetPathToExeDirectory()
        {
            string path_exe_directory = Application.StartupPath;

            return path_exe_directory;
        }

        /// <summary>Get full name to the config (XML) file for this application.</summary>
        public static string ConfigFileName()
        {
            string config_file_name = @"SendNewsletter.config";

            string config_file_path = FileUtil.GetPathToExeDirectory();

            string full_name = Path.Combine(config_file_path, config_file_name);

            return full_name;
        }

        /// <summary>Get the name to the poster directory. Create the directory if not existing.</summary>
        public static string PosterDirectory()
        {
            string exe_dir_path = FileUtil.GetPathToExeDirectory();

            string dir_poster = SendNewsLetter.NewsLetterSettings.Default.DirPoster;

            string poster_directory = Path.Combine(exe_dir_path, dir_poster);

            if (!Directory.Exists(poster_directory))
            {
                Directory.CreateDirectory(poster_directory);
            }

            return poster_directory;
        }

        /// <summary>Get the name to the new version directory. Create the directory if not existing.</summary>
        public static string NewVersionDirectory()
        {
            string exe_dir_path = FileUtil.GetPathToExeDirectory();

            string dir_new_version = SendNewsLetter.NewsLetterSettings.Default.DirNewVersion;

            string new_version_directory = Path.Combine(exe_dir_path, dir_new_version);

            if (!Directory.Exists(new_version_directory))
            {
                Directory.CreateDirectory(new_version_directory);
            }

            return new_version_directory;
        }

        /// <summary>Get the name to the templates directory. Create the directory if not existing.</summary>
        public static string TemplatesDirectory()
        {
            string exe_dir_path = FileUtil.GetPathToExeDirectory();

            string dir_templates = SendNewsLetter.NewsLetterSettings.Default.DirTemplates;

            string templates_directory = Path.Combine(exe_dir_path, dir_templates);

            if (!Directory.Exists(templates_directory))
            {
                Directory.CreateDirectory(templates_directory);
            }

            return templates_directory;
        }

        /// <summary>Get the name to the help directory. Create the directory if not existing.</summary>
        public static string HelpDirectory()
        {
            string exe_dir_path = FileUtil.GetPathToExeDirectory();

            string dir_help = SendNewsLetter.NewsLetterSettings.Default.DirHelp;

            string help_directory = Path.Combine(exe_dir_path, dir_help);

            if (!Directory.Exists(help_directory))
            {
                Directory.CreateDirectory(help_directory);
            }

            return help_directory;
        }

        /// <summary>Get full name to the help file.</summary>
        public static string HelpFileName()
        {
            string file_name_help = SendNewsLetter.NewsLetterSettings.Default.FileHelp;

            string path_file_name_help = Path.Combine(HelpDirectory(), file_name_help);

            return path_file_name_help;
        }


        /// <summary>Create the file if missing, i.e. copy from the resorces to the help directory.</summary>
        public static void HelpFileCreateIfMissing()
        {
            if (File.Exists(HelpFileName()))
            {
                return; // File exists already. Do nothing.
            }

            string help_file_resources = Properties.Resources.JAZZ_live_AARAU_Newsletter;

            string o_error = "";

            try
            {
                using (FileStream fileStream = new FileStream(HelpFileName(), FileMode.Create))
                // Without System.Text.Encoding.Default there are problems with ä ö ü
                using (StreamWriter stream_writer = new StreamWriter(fileStream, System.Text.Encoding.Default))
                {
                    stream_writer.Write(help_file_resources);

                    stream_writer.Close();
                }
            }

            catch (FileNotFoundException) { o_error = "File not found"; return; }
            catch (DirectoryNotFoundException) { o_error = "Directory not found"; return; }
            catch (InvalidOperationException) { o_error = "Invalid operation"; return; }
            catch (InvalidCastException) { o_error = "invalid cast"; return; }
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return;
            }
        }

        /// <summary>Get the name to the attachment directory. Create the directory if not existing.</summary>
        public static string AttachmentDirectory()
        {
            string exe_dir_path = FileUtil.GetPathToExeDirectory();

            string dir_attachment = SendNewsLetter.NewsLetterSettings.Default.DirAttachment;

            string attachment_directory = Path.Combine(exe_dir_path, dir_attachment);

            if (!Directory.Exists(attachment_directory))
            {
                Directory.CreateDirectory(attachment_directory);
            }

            return attachment_directory;
        }

        /// <summary>Get full name to the Email adresses directory. Create the directory if not existing.</summary>
        public static string EmailAdressesDirectory()
        {
            string exe_dir_path = FileUtil.GetPathToExeDirectory();

            string dir_email_addresses = SendNewsLetter.NewsLetterSettings.Default.DirExcel;

            string email_addresses_directory = Path.Combine(exe_dir_path, dir_email_addresses);

            if (!Directory.Exists(email_addresses_directory))
            {
                Directory.CreateDirectory(email_addresses_directory);
            }

            return email_addresses_directory;
        }

        /// <summary>Get full name to the Excel file with Email addresses.</summary>
        public static string InputCsvFileName()
        {
            string file_name_csv = SendNewsLetter.NewsLetterSettings.Default.FileExcel;

            string path_file_name_csv = Path.Combine(EmailAdressesDirectory(), file_name_csv);

            return path_file_name_csv;
        }

        /// <summary>
        /// Get files with given extensions
        /// </summary>
        /// <param name="i_extensions">Array of extensions (with point)</param>
        /// <param name="i_directory">Search directory</param>
        /// <param name="o_file_names">Array of found files with paths</param>
        /// <returns>false if directory not exists or the input array of extensions is empty</returns>
        static public bool GetFilesDirectory(string[] i_extensions, string i_directory, out string[] o_file_names)
        {
            ArrayList files_string_array = new ArrayList();
            o_file_names = (string[])files_string_array.ToArray(typeof(string));

            for (int i_ext = 0; i_ext < i_extensions.Length; i_ext++)
            {
                string current_ext = i_extensions[i_ext];

                string[] files_ext = Directory.GetFiles(i_directory, "*" + current_ext);

                foreach (string file_ext in files_ext)
                {
                    files_string_array.Add(file_ext);
                }
            }

            files_string_array.Reverse();

            o_file_names = (string[])files_string_array.ToArray(typeof(string));

            return true;
        }


        /// <summary>Returns true if the file is in the array.</summary>
        public static bool FileIsInArray(string i_file_name, string[] i_file_names_array)
        {
            foreach (string file_name_array in i_file_names_array )
            {
                if (String.Compare(i_file_name, file_name_array, false) == 0)
                {
                    return true;
                }

            }
 
            return false;
        } // FileIsInArray

    } // Class FileUtil

    /// <summary>Mail (for reservation) utility functions</summary>
    public static class MailUtil
    {
        /// <summary>New line for windows</summary>
        public static string NewLineWindows = "\r\n";

        /// <summary>New line for unix</summary>
        public static string NewLineUnix = "\n";

        /// <summary>New line for mailto (body)</summary>
        public static string NewLineMail = "%0D%0A";

        /// <summary>Space</summary>
        public static string Space = "%20";

        /// <summary>Change text space to %20</summary>
        public static string SpaceToPercentage(string i_string)
        {
            string r_html_text_space_mod = "";
            if (i_string.Length == 0)
                return r_html_text_space_mod;

            for (int i_char = 0; i_char < i_string.Length; i_char++)
            {
                string one_char = i_string.Substring(i_char, 1);
                if (one_char == " ")
                    r_html_text_space_mod = r_html_text_space_mod + Space;
                else
                    r_html_text_space_mod = r_html_text_space_mod + one_char;
            }

            return r_html_text_space_mod;
        }

        /// <summary>Returns new lines for a mailto body (text)</summary>
        public static string NewLines(string i_body_text, bool i_b_proc_20)
        {
            string ret_str = @"";

            if (i_b_proc_20)
            {
                ret_str = ret_str + SpaceToPercentage(i_body_text);
            }

            bool b_new_line_unix = false;
            bool b_new_line_windows = i_body_text.Contains(NewLineWindows);
            if (b_new_line_windows)
            {
                b_new_line_unix = true; // Not used
            }
            else
            {
                b_new_line_unix = i_body_text.Contains(NewLineUnix);
            }

            if (!b_new_line_windows && !b_new_line_unix)
            {
                return ret_str;
            }

            if (b_new_line_windows)
            {
                ret_str = ret_str.Replace(NewLineWindows, NewLineMail);
                return ret_str;
            }

            ret_str = ret_str.Replace(NewLineUnix, NewLineMail);

            return ret_str;

        } // NewLines


    } // MailUtil

    /// <summary>Tooltip utility functions</summary>
    public static class ToolTipUtil
    {
        /// <summary>Set the time the tool tip shall be shown, how quick, etc.</summary>
        public static void SetDelays(ref ToolTip io_tool_tip)
        {
            // Set up the delays for the ToolTip.
            io_tool_tip.AutoPopDelay = 50000; // Default is 5000
            io_tool_tip.InitialDelay = 500;  // Default is 1000
            io_tool_tip.ReshowDelay = 100; // Default is 500
            // Force the ToolTip text to be displayed whether or not the form is active.
            io_tool_tip.ShowAlways = true;

        }


    }

    /// <summary>Time utility functions</summary>
    public static class TimeUtil
    {
        /// <summary>Returns minute, second and millisecond as a string</summary>
        public static string MinSecMilli()
        {
            DateTime current_time = DateTime.Now;
            int now_minute = current_time.Minute;
            int now_second = current_time.Second;
            int now_millisecond = current_time.Millisecond;

            string time_text = now_minute.ToString() + ":" + now_second.ToString() + ":" + now_millisecond.ToString();

            return time_text;
        }


        /// <summary>Returns date and time as a string</summary>
        public static string YearMonthDayHourMinSec()
        {
            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;
            int now_day = current_time.Day;
            int now_hour = current_time.Hour;
            int now_minute = current_time.Minute;
            int now_second = current_time.Second;

            string time_text = now_year.ToString() + "-" + _IntToString(now_month) + "-" + _IntToString(now_day) + "  " + _IntToString(now_hour) + ":" + _IntToString(now_minute) + ":" + _IntToString(now_second);

            return time_text;

        } // YearMonthDayHourMinSec

        /// <summary>Returns date + time + computer as a string</summary>
        public static string YearMonthDayHourMinSecComputer()
        {

            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;
            int now_day = current_time.Day;
            int now_hour = current_time.Hour;
            int now_minute = current_time.Minute;
            int now_second = current_time.Second;

            string time_text = now_year.ToString() + "_" + _IntToString(now_month) + "_" + _IntToString(now_day) + "_" + _IntToString(now_hour) + "_" + _IntToString(now_minute) + "_" + _IntToString(now_second);

            string time_text_computer = time_text + @"_" + System.Environment.MachineName;

            return time_text_computer;

        } // YearMonthDayHourMinSecComputer

        /// <summary>Returns current year, mont and day as an int array</summary>
        public static int[] YearMonthDayArray()
        {
            int[] ret_date_array = new int[3];

            DateTime current_time = DateTime.Now;

            ret_date_array[0] = current_time.Year;

            ret_date_array[1] = current_time.Month;

            ret_date_array[2] = current_time.Day;

            return ret_date_array;

        } // YearMonthDayArray

        /// <summary>Returns date and time as a string</summary>
        private static string _IntToString(int i_int)
        {
            string time_text = i_int.ToString();

            if (i_int <= 9)
            {
                time_text = "0" + time_text;
            }

            return time_text;
        }
    } // Class TimeUtil

    /// <summary>Functions that handle the seasons. Copied from application AddressesJazz. TODO Make it to an assembly</summary>
    static public class Season
    {
        /// <summary>Start month for a new season is four (April). After the first of April a new season .csv will be created by the old application.</summary>
        static private int m_start_date_new_season = 4; // April 

        /// <summary>Returns the current season as string. </summary>
        static public string GetCurrentSeason()
        {
            string ret_current_season = "";

            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;

            if (now_month < m_start_date_new_season)
            {
                ret_current_season = (now_year - 1).ToString() + "-" + now_year.ToString();
            }
            else
            {
                ret_current_season = now_year.ToString() + "-" + (now_year + 1).ToString();
            }

            return ret_current_season;
        }
    } // Class Season

} // namespace
