using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml;
using System.Windows.Forms;

namespace SendNewsLetter
{
    /// <summary>Edit of the config file and update of the main form of the application</summary>
    class EditConfigFile
    {
        /// <summary>Config file name with path</summary>
        private string m_config_file_name;

        /// <summary>Main form</summary>
        private NewsletterForm m_main_form = null;

        /// <summary>The names the XML elements in the config file</summary>
        private string[] m_config_file_names;

        /// <summary>The inner texts of the XML elements in the config file</summary>
        private string[] m_config_file_inner_texts;

        /// <summary>The resulting names from the search</summary>
        private string[] m_search_names;

        /// <summary>The resulting inner texts from the search</summary>
        private string[] m_search_inner_texts;

        /// <summary>Flag telling if XML object exists</summary>
        private bool m_load_xml = false;

        public bool UpdateConfigFile(string i_xml_element_name, string i_xml_inner_text, out string o_error)
        {
            o_error = "";

            XmlDocument xml_doc_modify = new XmlDocument();

            if (!_LoadXml(ref xml_doc_modify, out o_error))
            {
                return false;
            }

            string str_root_element = NewsLetterSettings.Default.ConfigRootElement;
            XmlNode nood_settings = xml_doc_modify.SelectSingleNode(str_root_element);
            if (null == nood_settings)
            {
                o_error = "UpdateConfigFile programming error. There is no node: " + str_root_element;
                return false;
            }

            XmlNodeList noods_input = nood_settings.ChildNodes;

            foreach (XmlNode nood_input in noods_input)
            {
                string xml_element_name = nood_input.Name;
                string xml_element_inner_text = nood_input.InnerText;

                if (i_xml_element_name == xml_element_name)
                {
                    nood_input.InnerText = i_xml_inner_text;
                    break;
                }
            }

            xml_doc_modify.Save(this.m_config_file_name);

            return true;
        }

        /// <summary>Get the elements that contains the input string</summary>
        public bool Search(string i_str_search, out string[] o_search_names, out string[] o_search_inner_texts)
        {
            ArrayList search_name_string_array = new ArrayList();
            ArrayList search_inner_text_string_array = new ArrayList();

            this.m_search_names = (string[])search_name_string_array.ToArray(typeof(string));
            this.m_search_inner_texts = (string[])search_inner_text_string_array.ToArray(typeof(string));
            o_search_names = this.m_search_names;
            o_search_inner_texts = this.m_search_inner_texts;

            if (i_str_search.Trim() == "")
            {
                return true;
            }

            string error_message;
            m_load_xml = _GetConfigLines(out error_message);
            if (!m_load_xml)
            {
                MessageBox.Show(error_message);
                return false;
            }

            string str_search_lower_case = i_str_search.ToLower(); 

            for (int i_line = 0; i_line < m_config_file_inner_texts.Length; i_line++)
            {
                string current_name = m_config_file_names[i_line];
                string current_innertext = m_config_file_inner_texts[i_line];
                string current_innertext_lower_case = current_innertext.ToLower();

                int index_substring = current_innertext_lower_case.IndexOf(str_search_lower_case);
                if (index_substring >= 0)
                {
                    search_name_string_array.Add(current_name);
                    search_inner_text_string_array.Add(current_innertext);
                }
            }

            this.m_search_names = (string[])search_name_string_array.ToArray(typeof(string));
            this.m_search_inner_texts = (string[])search_inner_text_string_array.ToArray(typeof(string));
            o_search_names = this.m_search_names;
            o_search_inner_texts = this.m_search_inner_texts;

            return true;
        }

        /// <summary>Constructor setting the config file name and the main form</summary>
        public EditConfigFile(string i_config_file_name, NewsletterForm i_main_form)
        {
            m_config_file_name = i_config_file_name;
            m_main_form = i_main_form;
        }


        /// <summary>Constructor setting the config file name and the main form</summary>
        private bool _GetConfigLines(out string o_error)
        {
            o_error = "";

            ArrayList config_name_string_array = new ArrayList();
            ArrayList config_inner_text_string_array = new ArrayList();

            this.m_config_file_names = (string[])config_name_string_array.ToArray(typeof(string));
            this.m_config_file_inner_texts = (string[])config_inner_text_string_array.ToArray(typeof(string));

            XmlDocument xml_doc_input = new XmlDocument();

            if (!_LoadXml(ref xml_doc_input, out o_error))
            {
                return false;
            }

            string str_root_element = NewsLetterSettings.Default.ConfigRootElement;
            XmlNode nood_settings = xml_doc_input.SelectSingleNode(str_root_element);
            if (null == nood_settings)
            {
                o_error = "_GetConfigLines programming error. There is no node: " + str_root_element;
                return false;
            }

            XmlNodeList noods_input = nood_settings.ChildNodes;

            foreach (XmlNode nood_input in noods_input)
            {                
                string xml_element_name = nood_input.Name;
                string xml_element_inner_text = nood_input.InnerText;

                config_name_string_array.Add(xml_element_name);
                config_inner_text_string_array.Add(xml_element_inner_text);
            }

            this.m_config_file_names = (string[])config_name_string_array.ToArray(typeof(string));
            this.m_config_file_inner_texts = (string[])config_inner_text_string_array.ToArray(typeof(string));

            return true;
        }

        /// <summary>Load input config (XML) file. </summary>
        private bool _LoadXml(ref XmlDocument o_xml_doc_input, out string o_error)
        {
            o_error = "";

            try
            {
                o_xml_doc_input.Load(this.m_config_file_name);
            }
            catch (XmlException xmlEx)
            {
                o_error = xmlEx.Message;
                return false;
            }

            catch (Exception ex)
            {
                o_error = ex.Message;
                return false;
            }

            return true;
        }

    }
}
