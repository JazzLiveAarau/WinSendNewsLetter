using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SendNewsLetter
{
    /// <summary>Shows help</summary>
    public partial class HelpForm : Form
    {
        private ToolTip m_tool_tip_help = new ToolTip();
        private ToolTip m_tool_tip_help_text = new ToolTip();
        private ToolTip m_tool_tip_help_exit = new ToolTip();

        /// <summary>Constructor </summary>
        public HelpForm()
        {
            InitializeComponent();

            this.Text = NewsLetterSettings.Default.GuiHelpDialogTitle;

            this.m_button_help_exit.Text = NewsLetterSettings.Default.GuiHelpDialogExit;

            this.m_rich_text_box_help.LoadFile(FileUtil.HelpFileName(), RichTextBoxStreamType.RichText);

            m_tool_tip_help.SetToolTip(this, NewsLetterSettings.Default.ToolTipButtonHelp);
            ToolTipUtil.SetDelays(ref m_tool_tip_help);
            m_tool_tip_help_text.SetToolTip(this.m_rich_text_box_help, NewsLetterSettings.Default.ToolTipHelpTextBox);
            ToolTipUtil.SetDelays(ref m_tool_tip_help_text);
            m_tool_tip_help_exit.SetToolTip(this.m_button_help_exit, NewsLetterSettings.Default.ToolTipHelpExit);
            ToolTipUtil.SetDelays(ref m_tool_tip_help_exit);

        }

        /// <summary>Exit from help dialog</summary>
        private void m_button_help_exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
