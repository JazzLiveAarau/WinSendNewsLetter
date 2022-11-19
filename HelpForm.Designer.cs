namespace SendNewsLetter
{
    partial class HelpForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HelpForm));
            this.m_rich_text_box_help = new System.Windows.Forms.RichTextBox();
            this.m_button_help_exit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // m_rich_text_box_help
            // 
            this.m_rich_text_box_help.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.m_rich_text_box_help.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_rich_text_box_help.Location = new System.Drawing.Point(5, 7);
            this.m_rich_text_box_help.Name = "m_rich_text_box_help";
            this.m_rich_text_box_help.Size = new System.Drawing.Size(659, 253);
            this.m_rich_text_box_help.TabIndex = 0;
            this.m_rich_text_box_help.Text = "";
            // 
            // m_button_help_exit
            // 
            this.m_button_help_exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_help_exit.Location = new System.Drawing.Point(577, 269);
            this.m_button_help_exit.Name = "m_button_help_exit";
            this.m_button_help_exit.Size = new System.Drawing.Size(72, 22);
            this.m_button_help_exit.TabIndex = 1;
            this.m_button_help_exit.Text = "Exit";
            this.m_button_help_exit.UseVisualStyleBackColor = true;
            this.m_button_help_exit.Click += new System.EventHandler(this.m_button_help_exit_Click);
            // 
            // HelpForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(667, 293);
            this.Controls.Add(this.m_button_help_exit);
            this.Controls.Add(this.m_rich_text_box_help);
            this.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "HelpForm";
            this.ShowIcon = false;
            this.Text = "Help Newsletter";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox m_rich_text_box_help;
        private System.Windows.Forms.Button m_button_help_exit;
    }
}