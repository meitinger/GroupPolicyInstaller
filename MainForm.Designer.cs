namespace Aufbauwerk.Tools.GroupPolicyInstaller
{
    partial class MainForm
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.logoPictureBox = new System.Windows.Forms.PictureBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.productLabel = new System.Windows.Forms.Label();
            this.actionLabel = new System.Windows.Forms.Label();
            this.table = new System.Windows.Forms.TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).BeginInit();
            this.table.SuspendLayout();
            this.SuspendLayout();
            // 
            // logoPictureBox
            // 
            this.table.SetColumnSpan(this.logoPictureBox, 3);
            this.logoPictureBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logoPictureBox.Location = new System.Drawing.Point(3, 3);
            this.logoPictureBox.Name = "logoPictureBox";
            this.logoPictureBox.Size = new System.Drawing.Size(610, 370);
            this.logoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.logoPictureBox.TabIndex = 0;
            this.logoPictureBox.TabStop = false;
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar.ForeColor = System.Drawing.Color.Silver;
            this.progressBar.Location = new System.Drawing.Point(111, 392);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(394, 12);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 1;
            this.progressBar.Value = 50;
            // 
            // productLabel
            // 
            this.productLabel.AutoEllipsis = true;
            this.productLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.productLabel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.productLabel.ForeColor = System.Drawing.Color.White;
            this.productLabel.Location = new System.Drawing.Point(108, 376);
            this.productLabel.Margin = new System.Windows.Forms.Padding(0);
            this.productLabel.Name = "productLabel";
            this.productLabel.Size = new System.Drawing.Size(400, 13);
            this.productLabel.TabIndex = 0;
            this.productLabel.Text = "<Product>";
            // 
            // actionLabel
            // 
            this.actionLabel.AutoEllipsis = true;
            this.actionLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.actionLabel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.actionLabel.ForeColor = System.Drawing.Color.White;
            this.actionLabel.Location = new System.Drawing.Point(108, 407);
            this.actionLabel.Margin = new System.Windows.Forms.Padding(0);
            this.actionLabel.Name = "actionLabel";
            this.actionLabel.Size = new System.Drawing.Size(400, 13);
            this.actionLabel.TabIndex = 2;
            this.actionLabel.Text = "<Action>";
            // 
            // table
            // 
            this.table.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.table.ColumnCount = 3;
            this.table.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.table.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 400F));
            this.table.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.table.Controls.Add(this.actionLabel, 1, 3);
            this.table.Controls.Add(this.productLabel, 1, 1);
            this.table.Controls.Add(this.progressBar, 1, 2);
            this.table.Controls.Add(this.logoPictureBox, 0, 0);
            this.table.Location = new System.Drawing.Point(12, 12);
            this.table.Name = "table";
            this.table.RowCount = 5;
            this.table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.table.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 18F));
            this.table.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.table.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.table.Size = new System.Drawing.Size(616, 456);
            this.table.TabIndex = 3;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(640, 480);
            this.Controls.Add(this.table);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "MainForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).EndInit();
            this.table.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox logoPictureBox;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label productLabel;
        private System.Windows.Forms.Label actionLabel;
        private System.Windows.Forms.TableLayoutPanel table;

    }
}

