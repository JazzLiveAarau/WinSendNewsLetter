using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace SendNewsLetter
{
    /// <summary>Form for edit of the config (XML) file of the application</summary>
    public partial class EditConfigFileForm : Form
    {
        private NewsletterForm m_main_form = null;

        /// <summary>Object for editing of the config file</summary>
        private EditConfigFile m_edit_config_file = null;

        /// <summary>The resulting inner texts from the search</summary>
        private string[] m_search_inner_texts;

        /// <summary>The resulting names from the search</summary>
        private string[] m_search_names;

        /// <summary>Index for the selected XML element</summary>
        private int m_selected_index = -1;

        /// <summary>Constructor creaing an instance of EditConfigFile</summary>
        public EditConfigFileForm(NewsletterForm i_main_form)
        {
            InitializeComponent();

            this.m_main_form = i_main_form;

            this.m_edit_config_file = new EditConfigFile(FileUtil.ConfigFileName(), this.m_main_form);

            this.Text = "Edit of " + Path.GetFileName(FileUtil.ConfigFileName());
        }

        /// <summary>Search all XML elements that contains a given string</summary>
        private void m_button_search_Click(object sender, EventArgs e)
        {
            _Search();
        }

        /// <summary>Search all XML elements that contains a given string</summary>
        private void _Search()
        {
            this.m_edit_config_file.Search(m_text_box_search.Text, out m_search_names, out m_search_inner_texts);

            this.m_combo_box_xml_elements.Text = "";
            this.m_text_box_xml_start.Text = "";
            this.m_text_box_edit.Text = "";

            this.m_text_box_number_hits.Text = m_search_inner_texts.Length.ToString();

            this.m_combo_box_xml_elements.Items.Clear();

            for (int i_line = 0; i_line < m_search_inner_texts.Length; i_line++)
            {
                string current_inner_text = m_search_inner_texts[i_line];

                string current_name = m_search_names[i_line];

                this.m_combo_box_xml_elements.Items.Add(current_inner_text);

                if (0 == i_line)
                {
                    this.m_combo_box_xml_elements.Text = current_inner_text;
                    this.m_text_box_xml_start.Text = "<" + current_name + ">";
                    this.m_selected_index = 0;
                }
            }
        }

        /// <summary>Update the configuration file</summary>
        private void m_button_save_Click(object sender, EventArgs e)
        {
            if (this.m_text_box_edit.Text.Trim() == "")
            {
                return;
            }

            string save_inner_text = "";
            string save_name = "";
            if (this.m_selected_index >= 0 && this.m_selected_index < m_search_inner_texts.Length)
            {
                save_inner_text = this.m_text_box_edit.Text;
                save_name = m_search_names[this.m_selected_index];
            }
            else
            {
                MessageBox.Show("Programming error m_button_save_Click Index =" + this.m_selected_index.ToString());
                return;
            }

            string error_message = "";
            if (!this.m_edit_config_file.UpdateConfigFile(save_name, save_inner_text, out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            this.m_text_box_search.Text = "";
            this.m_combo_box_xml_elements.Text = "";
            this.m_text_box_xml_start.Text = "";
            this.m_text_box_edit.Text = "";
            this.m_combo_box_xml_elements.Items.Clear();
            this.m_selected_index = -1;

            this.m_main_form.ConfigFileHasBeenChanged();
        }

        private void m_button_exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void m_combo_box_xml_elements_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selected_inner_text = this.m_combo_box_xml_elements.Text;
            this.m_selected_index = this.m_combo_box_xml_elements.SelectedIndex;
            string selected_name = m_search_names[this.m_selected_index];
            this.m_text_box_xml_start.Text = "<" + selected_name + ">";
            this.m_text_box_edit.Text = selected_inner_text;
        }

        private void m_text_box_search_TextChanged(object sender, EventArgs e)
        {
            _Search();
        }
    }
}
