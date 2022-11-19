namespace SendNewsLetter
{
    partial class EditConfigFileForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.m_text_box_search = new System.Windows.Forms.TextBox();
            this.m_combo_box_xml_elements = new System.Windows.Forms.ComboBox();
            this.m_text_box_edit = new System.Windows.Forms.TextBox();
            this.m_button_save = new System.Windows.Forms.Button();
            this.m_button_exit = new System.Windows.Forms.Button();
            this.m_text_box_xml_start = new System.Windows.Forms.TextBox();
            this.m_text_box_number_hits = new System.Windows.Forms.TextBox();
            this.m_label_search_string = new System.Windows.Forms.Label();
            this.m_label_xml_element = new System.Windows.Forms.Label();
            this.m_label_xml_value = new System.Windows.Forms.Label();
            this.m_label_number_hits = new System.Windows.Forms.Label();
            this.m_label_edit = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // m_text_box_search
            // 
            this.m_text_box_search.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.m_text_box_search.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_text_box_search.Location = new System.Drawing.Point(24, 47);
            this.m_text_box_search.Name = "m_text_box_search";
            this.m_text_box_search.Size = new System.Drawing.Size(307, 20);
            this.m_text_box_search.TabIndex = 0;
            this.m_text_box_search.TextChanged += new System.EventHandler(this.m_text_box_search_TextChanged);
            // 
            // m_combo_box_xml_elements
            // 
            this.m_combo_box_xml_elements.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.m_combo_box_xml_elements.BackColor = System.Drawing.SystemColors.MenuBar;
            this.m_combo_box_xml_elements.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_xml_elements.FormattingEnabled = true;
            this.m_combo_box_xml_elements.Location = new System.Drawing.Point(64, 130);
            this.m_combo_box_xml_elements.Name = "m_combo_box_xml_elements";
            this.m_combo_box_xml_elements.Size = new System.Drawing.Size(267, 22);
            this.m_combo_box_xml_elements.TabIndex = 2;
            this.m_combo_box_xml_elements.SelectedIndexChanged += new System.EventHandler(this.m_combo_box_xml_elements_SelectedIndexChanged);
            // 
            // m_text_box_edit
            // 
            this.m_text_box_edit.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.m_text_box_edit.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_text_box_edit.Location = new System.Drawing.Point(24, 174);
            this.m_text_box_edit.Name = "m_text_box_edit";
            this.m_text_box_edit.Size = new System.Drawing.Size(307, 20);
            this.m_text_box_edit.TabIndex = 3;
            // 
            // m_button_save
            // 
            this.m_button_save.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_save.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_save.Location = new System.Drawing.Point(280, 206);
            this.m_button_save.Name = "m_button_save";
            this.m_button_save.Size = new System.Drawing.Size(52, 26);
            this.m_button_save.TabIndex = 4;
            this.m_button_save.Text = "Save";
            this.m_button_save.UseVisualStyleBackColor = true;
            this.m_button_save.Click += new System.EventHandler(this.m_button_save_Click);
            // 
            // m_button_exit
            // 
            this.m_button_exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_exit.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_exit.Location = new System.Drawing.Point(281, 10);
            this.m_button_exit.Name = "m_button_exit";
            this.m_button_exit.Size = new System.Drawing.Size(52, 26);
            this.m_button_exit.TabIndex = 5;
            this.m_button_exit.Text = "Exit";
            this.m_button_exit.UseVisualStyleBackColor = true;
            this.m_button_exit.Click += new System.EventHandler(this.m_button_exit_Click);
            // 
            // m_text_box_xml_start
            // 
            this.m_text_box_xml_start.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.m_text_box_xml_start.Enabled = false;
            this.m_text_box_xml_start.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_text_box_xml_start.Location = new System.Drawing.Point(24, 91);
            this.m_text_box_xml_start.Name = "m_text_box_xml_start";
            this.m_text_box_xml_start.Size = new System.Drawing.Size(199, 20);
            this.m_text_box_xml_start.TabIndex = 6;
            // 
            // m_text_box_number_hits
            // 
            this.m_text_box_number_hits.Enabled = false;
            this.m_text_box_number_hits.Location = new System.Drawing.Point(24, 130);
            this.m_text_box_number_hits.Name = "m_text_box_number_hits";
            this.m_text_box_number_hits.Size = new System.Drawing.Size(34, 20);
            this.m_text_box_number_hits.TabIndex = 8;
            // 
            // m_label_search_string
            // 
            this.m_label_search_string.AutoSize = true;
            this.m_label_search_string.Location = new System.Drawing.Point(30, 29);
            this.m_label_search_string.Name = "m_label_search_string";
            this.m_label_search_string.Size = new System.Drawing.Size(72, 14);
            this.m_label_search_string.TabIndex = 9;
            this.m_label_search_string.Text = "Search string";
            // 
            // m_label_xml_element
            // 
            this.m_label_xml_element.AutoSize = true;
            this.m_label_xml_element.Location = new System.Drawing.Point(30, 73);
            this.m_label_xml_element.Name = "m_label_xml_element";
            this.m_label_xml_element.Size = new System.Drawing.Size(68, 14);
            this.m_label_xml_element.TabIndex = 10;
            this.m_label_xml_element.Text = "XML element";
            // 
            // m_label_xml_value
            // 
            this.m_label_xml_value.AutoSize = true;
            this.m_label_xml_value.Location = new System.Drawing.Point(71, 113);
            this.m_label_xml_value.Name = "m_label_xml_value";
            this.m_label_xml_value.Size = new System.Drawing.Size(101, 14);
            this.m_label_xml_value.TabIndex = 11;
            this.m_label_xml_value.Text = "Select XML element";
            // 
            // m_label_number_hits
            // 
            this.m_label_number_hits.AutoSize = true;
            this.m_label_number_hits.Location = new System.Drawing.Point(26, 113);
            this.m_label_number_hits.Name = "m_label_number_hits";
            this.m_label_number_hits.Size = new System.Drawing.Size(25, 14);
            this.m_label_number_hits.TabIndex = 12;
            this.m_label_number_hits.Text = "Hits";
            // 
            // m_label_edit
            // 
            this.m_label_edit.AutoSize = true;
            this.m_label_edit.Location = new System.Drawing.Point(32, 156);
            this.m_label_edit.Name = "m_label_edit";
            this.m_label_edit.Size = new System.Drawing.Size(117, 14);
            this.m_label_edit.TabIndex = 13;
            this.m_label_edit.Text = "Edit XML element value";
            // 
            // EditConfigFileForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(344, 244);
            this.Controls.Add(this.m_label_edit);
            this.Controls.Add(this.m_label_number_hits);
            this.Controls.Add(this.m_label_xml_value);
            this.Controls.Add(this.m_label_xml_element);
            this.Controls.Add(this.m_label_search_string);
            this.Controls.Add(this.m_text_box_number_hits);
            this.Controls.Add(this.m_text_box_xml_start);
            this.Controls.Add(this.m_button_exit);
            this.Controls.Add(this.m_button_save);
            this.Controls.Add(this.m_text_box_edit);
            this.Controls.Add(this.m_combo_box_xml_elements);
            this.Controls.Add(this.m_text_box_search);
            this.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "EditConfigFileForm";
            this.Text = "Edit of the configuration file";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox m_text_box_search;
        private System.Windows.Forms.ComboBox m_combo_box_xml_elements;
        private System.Windows.Forms.TextBox m_text_box_edit;
        private System.Windows.Forms.Button m_button_save;
        private System.Windows.Forms.Button m_button_exit;
        private System.Windows.Forms.TextBox m_text_box_xml_start;
        private System.Windows.Forms.TextBox m_text_box_number_hits;
        private System.Windows.Forms.Label m_label_search_string;
        private System.Windows.Forms.Label m_label_xml_element;
        private System.Windows.Forms.Label m_label_xml_value;
        private System.Windows.Forms.Label m_label_number_hits;
        private System.Windows.Forms.Label m_label_edit;
    }
}