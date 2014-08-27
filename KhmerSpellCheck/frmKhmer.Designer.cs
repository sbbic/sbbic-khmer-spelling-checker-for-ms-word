namespace KhmerSpellCheck
{
    partial class frmKhmer
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
            this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
            this.lblNotInDictionary = new System.Windows.Forms.Label();
            this.btnIgnore = new System.Windows.Forms.Button();
            this.btnIgnoreAll = new System.Windows.Forms.Button();
            this.lblSuggestions = new System.Windows.Forms.Label();
            this.btnChange = new System.Windows.Forms.Button();
            this.btnChangeAll = new System.Windows.Forms.Button();
            this.lstSuggestions = new System.Windows.Forms.ListBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSkip = new System.Windows.Forms.Button();
            this.tlpMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tlpMain
            // 
            this.tlpMain.AutoSize = true;
            this.tlpMain.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpMain.ColumnCount = 2;
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpMain.Controls.Add(this.lblNotInDictionary, 0, 0);
            this.tlpMain.Controls.Add(this.btnIgnore, 1, 1);
            this.tlpMain.Controls.Add(this.btnIgnoreAll, 1, 2);
            this.tlpMain.Controls.Add(this.lblSuggestions, 0, 4);
            this.tlpMain.Controls.Add(this.btnChange, 1, 5);
            this.tlpMain.Controls.Add(this.btnChangeAll, 1, 6);
            this.tlpMain.Controls.Add(this.lstSuggestions, 0, 5);
            this.tlpMain.Controls.Add(this.btnCancel, 1, 7);
            this.tlpMain.Controls.Add(this.btnSkip, 1, 3);
            this.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpMain.Location = new System.Drawing.Point(0, 0);
            this.tlpMain.Name = "tlpMain";
            this.tlpMain.Padding = new System.Windows.Forms.Padding(5);
            this.tlpMain.RowCount = 8;
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.Size = new System.Drawing.Size(505, 176);
            this.tlpMain.TabIndex = 0;
            // 
            // lblNotInDictionary
            // 
            this.lblNotInDictionary.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblNotInDictionary.AutoSize = true;
            this.tlpMain.SetColumnSpan(this.lblNotInDictionary, 2);
            this.lblNotInDictionary.Location = new System.Drawing.Point(8, 5);
            this.lblNotInDictionary.Name = "lblNotInDictionary";
            this.lblNotInDictionary.Size = new System.Drawing.Size(91, 13);
            this.lblNotInDictionary.TabIndex = 0;
            this.lblNotInDictionary.Text = "Not in Dictionary :";
            this.lblNotInDictionary.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnIgnore
            // 
            this.btnIgnore.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnIgnore.AutoSize = true;
            this.btnIgnore.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnIgnore.DialogResult = System.Windows.Forms.DialogResult.Ignore;
            this.btnIgnore.Location = new System.Drawing.Point(397, 21);
            this.btnIgnore.MinimumSize = new System.Drawing.Size(100, 0);
            this.btnIgnore.Name = "btnIgnore";
            this.btnIgnore.Size = new System.Drawing.Size(100, 23);
            this.btnIgnore.TabIndex = 1;
            this.btnIgnore.Text = "Ignore";
            this.btnIgnore.UseVisualStyleBackColor = true;
            // 
            // btnIgnoreAll
            // 
            this.btnIgnoreAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnIgnoreAll.AutoSize = true;
            this.btnIgnoreAll.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnIgnoreAll.Location = new System.Drawing.Point(397, 50);
            this.btnIgnoreAll.MinimumSize = new System.Drawing.Size(100, 0);
            this.btnIgnoreAll.Name = "btnIgnoreAll";
            this.btnIgnoreAll.Size = new System.Drawing.Size(100, 23);
            this.btnIgnoreAll.TabIndex = 2;
            this.btnIgnoreAll.Text = "Ignore All";
            this.btnIgnoreAll.UseVisualStyleBackColor = true;
            this.btnIgnoreAll.Click += new System.EventHandler(this.btnIgnoreAll_Click);
            // 
            // lblSuggestions
            // 
            this.lblSuggestions.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblSuggestions.AutoSize = true;
            this.tlpMain.SetColumnSpan(this.lblSuggestions, 2);
            this.lblSuggestions.Location = new System.Drawing.Point(8, 105);
            this.lblSuggestions.Name = "lblSuggestions";
            this.lblSuggestions.Size = new System.Drawing.Size(71, 13);
            this.lblSuggestions.TabIndex = 5;
            this.lblSuggestions.Text = "Suggestions :";
            this.lblSuggestions.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnChange
            // 
            this.btnChange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChange.AutoSize = true;
            this.btnChange.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnChange.Location = new System.Drawing.Point(397, 121);
            this.btnChange.MinimumSize = new System.Drawing.Size(100, 0);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(100, 23);
            this.btnChange.TabIndex = 7;
            this.btnChange.Text = "Change";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // btnChangeAll
            // 
            this.btnChangeAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChangeAll.AutoSize = true;
            this.btnChangeAll.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnChangeAll.Location = new System.Drawing.Point(397, 150);
            this.btnChangeAll.MinimumSize = new System.Drawing.Size(100, 0);
            this.btnChangeAll.Name = "btnChangeAll";
            this.btnChangeAll.Size = new System.Drawing.Size(100, 23);
            this.btnChangeAll.TabIndex = 8;
            this.btnChangeAll.Text = "Change All";
            this.btnChangeAll.UseVisualStyleBackColor = true;
            this.btnChangeAll.Click += new System.EventHandler(this.btnChangeAll_Click);
            // 
            // lstSuggestions
            // 
            this.lstSuggestions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstSuggestions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstSuggestions.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstSuggestions.FormattingEnabled = true;
            this.lstSuggestions.ItemHeight = 24;
            this.lstSuggestions.Location = new System.Drawing.Point(8, 121);
            this.lstSuggestions.Name = "lstSuggestions";
            this.tlpMain.SetRowSpan(this.lstSuggestions, 3);
            this.lstSuggestions.Size = new System.Drawing.Size(383, 81);
            this.lstSuggestions.TabIndex = 12;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.AutoSize = true;
            this.btnCancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(397, 179);
            this.btnCancel.MinimumSize = new System.Drawing.Size(100, 0);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 23);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnSkip
            // 
            this.btnSkip.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSkip.AutoSize = true;
            this.btnSkip.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnSkip.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.btnSkip.Location = new System.Drawing.Point(397, 79);
            this.btnSkip.Name = "btnSkip";
            this.btnSkip.Size = new System.Drawing.Size(100, 23);
            this.btnSkip.TabIndex = 13;
            this.btnSkip.Text = "Skip && Edit";
            this.btnSkip.UseVisualStyleBackColor = true;
            // 
            // frmKhmer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(505, 176);
            this.Controls.Add(this.tlpMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmKhmer";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Spelling : Khmer";
            this.Load += new System.EventHandler(this.frmKhmer_Load);
            this.tlpMain.ResumeLayout(false);
            this.tlpMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tlpMain;
        private System.Windows.Forms.Label lblNotInDictionary;
        private System.Windows.Forms.Button btnIgnore;
        private System.Windows.Forms.Button btnIgnoreAll;
        private System.Windows.Forms.Label lblSuggestions;
        private System.Windows.Forms.Button btnChange;
        private System.Windows.Forms.Button btnChangeAll;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ListBox lstSuggestions;
        private System.Windows.Forms.Button btnSkip;
    }
}