namespace VintedCompanion
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnProcess = new Button();
            btnExportExcel = new Button();
            statusLabel = new Label();
            SuspendLayout();
            // 
            // btnProcess
            // 
            btnProcess.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnProcess.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btnProcess.Location = new Point(41, 41);
            btnProcess.Margin = new Padding(32);
            btnProcess.Name = "btnProcess";
            btnProcess.Size = new Size(718, 170);
            btnProcess.TabIndex = 0;
            btnProcess.Text = "Wybierz plik *.html z Vinted";
            btnProcess.UseVisualStyleBackColor = true;
            btnProcess.Click += btnProcess_Click;
            // 
            // btnExportExcel
            // 
            btnExportExcel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            btnExportExcel.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btnExportExcel.Location = new Point(41, 251);
            btnExportExcel.Margin = new Padding(32);
            btnExportExcel.Name = "btnExportExcel";
            btnExportExcel.Size = new Size(718, 170);
            btnExportExcel.TabIndex = 1;
            btnExportExcel.Text = "Wygeneruj plik Excel *.xlsx";
            btnExportExcel.UseVisualStyleBackColor = true;
            btnExportExcel.Click += btnExportExcel_Click;
            // 
            // statusLabel
            // 
            statusLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            statusLabel.AutoSize = true;
            statusLabel.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 0);
            statusLabel.Location = new Point(12, 443);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(69, 28);
            statusLabel.TabIndex = 2;
            statusLabel.Text = "Status:";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 480);
            Controls.Add(statusLabel);
            Controls.Add(btnExportExcel);
            Controls.Add(btnProcess);
            Name = "Form1";
            Text = "Vinted Generator";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnProcess;
        private Button btnExportExcel;
        private Label statusLabel;
    }
}
