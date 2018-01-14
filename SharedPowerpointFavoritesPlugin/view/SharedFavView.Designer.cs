namespace SharedPowerpointFavoritesPlugin
{
    partial class SharedFavView
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
            this.saveShapeButton = new System.Windows.Forms.Button();
            this.importButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.SuspendLayout();
            // 
            // saveShapeButton
            // 
            this.saveShapeButton.Location = new System.Drawing.Point(173, 367);
            this.saveShapeButton.Name = "saveShapeButton";
            this.saveShapeButton.Size = new System.Drawing.Size(89, 48);
            this.saveShapeButton.TabIndex = 1;
            this.saveShapeButton.Text = "Save From Clipboard";
            this.saveShapeButton.UseVisualStyleBackColor = true;
            this.saveShapeButton.Click += new System.EventHandler(this.saveShapeButton_Click);
            // 
            // importButton
            // 
            this.importButton.Location = new System.Drawing.Point(13, 366);
            this.importButton.Name = "importButton";
            this.importButton.Size = new System.Drawing.Size(77, 49);
            this.importButton.TabIndex = 4;
            this.importButton.Text = "Import...";
            this.importButton.UseVisualStyleBackColor = true;
            this.importButton.Click += new System.EventHandler(this.importButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.Location = new System.Drawing.Point(96, 367);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(71, 48);
            this.exportButton.TabIndex = 5;
            this.exportButton.Text = "Export...";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Location = new System.Drawing.Point(13, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(249, 348);
            this.tabControl1.TabIndex = 6;
            // 
            // SharedFavView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(274, 427);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.importButton);
            this.Controls.Add(this.saveShapeButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "SharedFavView";
            this.Text = "Shared FavoriteShapes";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SharedFavView_FormClosed);
            this.Load += new System.EventHandler(this.SharedFavView_Load);
            this.Shown += new System.EventHandler(this.SharedFavView_Shown);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button saveShapeButton;
        private System.Windows.Forms.Button importButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.TabControl tabControl1;
    }
}