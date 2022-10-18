namespace DataUnwrapping
{
    partial class DWColFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DWColFrm));
            this.GrBox = new System.Windows.Forms.GroupBox();
            this.CParList = new System.Windows.Forms.CheckedListBox();
            this.BtnOpen = new System.Windows.Forms.Button();
            this.SelectAllBtn = new System.Windows.Forms.Button();
            this.ClearBtn = new System.Windows.Forms.Button();
            this.GrBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // GrBox
            // 
            this.GrBox.BackColor = System.Drawing.SystemColors.Control;
            this.GrBox.Controls.Add(this.CParList);
            this.GrBox.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GrBox.ForeColor = System.Drawing.Color.Brown;
            this.GrBox.Location = new System.Drawing.Point(12, 12);
            this.GrBox.Name = "GrBox";
            this.GrBox.Size = new System.Drawing.Size(295, 350);
            this.GrBox.TabIndex = 0;
            this.GrBox.TabStop = false;
            this.GrBox.Text = "Choose Parameters";
            // 
            // CParList
            // 
            this.CParList.BackColor = System.Drawing.SystemColors.Window;
            this.CParList.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.CParList.CheckOnClick = true;
            this.CParList.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CParList.FormattingEnabled = true;
            this.CParList.Location = new System.Drawing.Point(5, 25);
            this.CParList.Margin = new System.Windows.Forms.Padding(5);
            this.CParList.Name = "CParList";
            this.CParList.Size = new System.Drawing.Size(286, 317);
            this.CParList.TabIndex = 0;
            // 
            // BtnOpen
            // 
            this.BtnOpen.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray;
            this.BtnOpen.Font = new System.Drawing.Font("Franklin Gothic Medium Cond", 12F);
            this.BtnOpen.ForeColor = System.Drawing.SystemColors.Highlight;
            this.BtnOpen.Location = new System.Drawing.Point(17, 376);
            this.BtnOpen.Name = "BtnOpen";
            this.BtnOpen.Size = new System.Drawing.Size(90, 31);
            this.BtnOpen.TabIndex = 1;
            this.BtnOpen.Text = "&Open Excel";
            this.BtnOpen.UseVisualStyleBackColor = true;
            this.BtnOpen.Click += new System.EventHandler(this.BtnOpen_Click);
            // 
            // SelectAllBtn
            // 
            this.SelectAllBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray;
            this.SelectAllBtn.Font = new System.Drawing.Font("Franklin Gothic Medium Cond", 12F);
            this.SelectAllBtn.ForeColor = System.Drawing.SystemColors.WindowFrame;
            this.SelectAllBtn.Location = new System.Drawing.Point(114, 376);
            this.SelectAllBtn.Name = "SelectAllBtn";
            this.SelectAllBtn.Size = new System.Drawing.Size(90, 31);
            this.SelectAllBtn.TabIndex = 2;
            this.SelectAllBtn.Text = "&Select All";
            this.SelectAllBtn.UseVisualStyleBackColor = true;
            this.SelectAllBtn.Click += new System.EventHandler(this.SelectAllBtn_Click);
            // 
            // ClearBtn
            // 
            this.ClearBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray;
            this.ClearBtn.Font = new System.Drawing.Font("Franklin Gothic Medium Cond", 12F);
            this.ClearBtn.ForeColor = System.Drawing.Color.OrangeRed;
            this.ClearBtn.Location = new System.Drawing.Point(213, 376);
            this.ClearBtn.Name = "ClearBtn";
            this.ClearBtn.Size = new System.Drawing.Size(90, 31);
            this.ClearBtn.TabIndex = 3;
            this.ClearBtn.Text = "&Clear";
            this.ClearBtn.UseVisualStyleBackColor = true;
            this.ClearBtn.Click += new System.EventHandler(this.ClearBtn_Click);
            // 
            // DWColFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(319, 423);
            this.Controls.Add(this.ClearBtn);
            this.Controls.Add(this.SelectAllBtn);
            this.Controls.Add(this.BtnOpen);
            this.Controls.Add(this.GrBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DWColFrm";
            this.Text = "Data Wrapper";
            this.GrBox.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox GrBox;
        private System.Windows.Forms.CheckedListBox CParList;
        private System.Windows.Forms.Button BtnOpen;
        private System.Windows.Forms.Button SelectAllBtn;
        private System.Windows.Forms.Button ClearBtn;
    }
}