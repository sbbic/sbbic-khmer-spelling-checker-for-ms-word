using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KhmerSpellCheck
{
    public partial class frmKhmer : Form
    {
        public bool ignoreAll = false;
        public string selectedSuggestion = string.Empty;
        public bool changeAll = false;

        public frmKhmer()
        {
            InitializeComponent();
        }

        public frmKhmer(string linetext, string khmerword, int startposition, List<String> suggestions, bool isObject)
        {
            InitializeComponent();

            tlpMain.Controls.Add(new RichTextBox
            {
                Name = "rtbSpellMistake",
                BackColor = Color.White,
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Enabled = true,
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            }, 0, 1);

            RichTextBox c = (RichTextBox)tlpMain.GetControlFromPosition(0, 1);
            c.Font = new System.Drawing.Font("Khmer UI", 12.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            tlpMain.SetRowSpan(c, 3);

            //Empty the suggestion box and text line box
            lstSuggestions.DataSource = null;

            //Set the line text box
            c.Text = linetext;
            c.Find(khmerword, startposition, RichTextBoxFinds.None);
            c.SelectionFont = new System.Drawing.Font("Khmer UI", 12.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            c.SelectionColor = Color.Red;
            c.ScrollToCaret();

            //Set the suggestion box
            lstSuggestions.DataSource = suggestions;

            if (isObject)
            {
                btnChange.Enabled = false;
            }
        }

        private void frmKhmer_Load(object sender, EventArgs e)
        {
            
        }

        private void btnIgnoreAll_Click(object sender, EventArgs e)
        {
            //Set ignore all to true
            ignoreAll = true;

            //Set the dialogresult
            this.DialogResult = System.Windows.Forms.DialogResult.Ignore;
        }

        private void btnChangeAll_Click(object sender, EventArgs e)
        {
            if (lstSuggestions.SelectedItem == null || String.IsNullOrWhiteSpace(lstSuggestions.SelectedItem.ToString()))
            {
                MessageBox.Show("Please select one of the provided suggestion(s).");
                return;
            }

            //Set the selected suggestion
            selectedSuggestion = lstSuggestions.SelectedItem.ToString();
            changeAll = true;

            //Set the dialogresult
            this.DialogResult = System.Windows.Forms.DialogResult.Yes;
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            if (lstSuggestions.SelectedItem == null || String.IsNullOrWhiteSpace(lstSuggestions.SelectedItem.ToString()))
            {
                MessageBox.Show("Please select one of the provided suggestion(s).");
                return;
            }

            //Set the selected suggestion
            selectedSuggestion = lstSuggestions.SelectedItem.ToString();

            //Set the dialogresult
            this.DialogResult = System.Windows.Forms.DialogResult.Yes;
        }
    }
}
