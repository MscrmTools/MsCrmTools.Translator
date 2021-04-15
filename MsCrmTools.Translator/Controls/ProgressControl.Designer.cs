namespace MsCrmTools.Translator.Controls
{
    public partial class ProgressControl
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlLeft = new System.Windows.Forms.Panel();
            this.pbProgress = new System.Windows.Forms.PictureBox();
            this.pnlMain = new System.Windows.Forms.Panel();
            this.lblError = new System.Windows.Forms.Label();
            this.lblSuccess = new System.Windows.Forms.Label();
            this.lblCount = new System.Windows.Forms.Label();
            this.pbTotal = new System.Windows.Forms.PictureBox();
            this.pbError = new System.Windows.Forms.PictureBox();
            this.pbSuccess = new System.Windows.Forms.PictureBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.pnlLeft.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbProgress)).BeginInit();
            this.pnlMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbTotal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbError)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbSuccess)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlLeft
            // 
            this.pnlLeft.Controls.Add(this.pbProgress);
            this.pnlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnlLeft.Location = new System.Drawing.Point(0, 0);
            this.pnlLeft.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pnlLeft.Name = "pnlLeft";
            this.pnlLeft.Size = new System.Drawing.Size(89, 82);
            this.pnlLeft.TabIndex = 0;
            // 
            // pbProgress
            // 
            this.pbProgress.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pbProgress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbProgress.Location = new System.Drawing.Point(0, 0);
            this.pbProgress.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pbProgress.Name = "pbProgress";
            this.pbProgress.Size = new System.Drawing.Size(89, 82);
            this.pbProgress.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pbProgress.TabIndex = 0;
            this.pbProgress.TabStop = false;
            // 
            // pnlMain
            // 
            this.pnlMain.Controls.Add(this.lblError);
            this.pnlMain.Controls.Add(this.lblSuccess);
            this.pnlMain.Controls.Add(this.lblCount);
            this.pnlMain.Controls.Add(this.pbTotal);
            this.pnlMain.Controls.Add(this.pbError);
            this.pnlMain.Controls.Add(this.pbSuccess);
            this.pnlMain.Controls.Add(this.lblTitle);
            this.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlMain.Location = new System.Drawing.Point(89, 0);
            this.pnlMain.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Size = new System.Drawing.Size(663, 82);
            this.pnlMain.TabIndex = 1;
            // 
            // lblError
            // 
            this.lblError.Location = new System.Drawing.Point(249, 34);
            this.lblError.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(75, 26);
            this.lblError.TabIndex = 6;
            this.lblError.Text = "lblError";
            this.lblError.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblSuccess
            // 
            this.lblSuccess.Location = new System.Drawing.Point(141, 34);
            this.lblSuccess.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblSuccess.Name = "lblSuccess";
            this.lblSuccess.Size = new System.Drawing.Size(75, 26);
            this.lblSuccess.TabIndex = 5;
            this.lblSuccess.Text = "lblSuccess";
            this.lblSuccess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCount
            // 
            this.lblCount.Location = new System.Drawing.Point(34, 34);
            this.lblCount.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblCount.Name = "lblCount";
            this.lblCount.Size = new System.Drawing.Size(75, 26);
            this.lblCount.TabIndex = 4;
            this.lblCount.Text = "lblCount";
            this.lblCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pbTotal
            // 
            this.pbTotal.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pbTotal.Image = global::MsCrmTools.Translator.Properties.Resources.sum;
            this.pbTotal.Location = new System.Drawing.Point(5, 34);
            this.pbTotal.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pbTotal.Name = "pbTotal";
            this.pbTotal.Size = new System.Drawing.Size(24, 26);
            this.pbTotal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbTotal.TabIndex = 3;
            this.pbTotal.TabStop = false;
            // 
            // pbError
            // 
            this.pbError.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pbError.Image = global::MsCrmTools.Translator.Properties.Resources.cross__1_;
            this.pbError.Location = new System.Drawing.Point(220, 34);
            this.pbError.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pbError.Name = "pbError";
            this.pbError.Size = new System.Drawing.Size(24, 26);
            this.pbError.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbError.TabIndex = 2;
            this.pbError.TabStop = false;
            // 
            // pbSuccess
            // 
            this.pbSuccess.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pbSuccess.Image = global::MsCrmTools.Translator.Properties.Resources.check;
            this.pbSuccess.Location = new System.Drawing.Point(113, 34);
            this.pbSuccess.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pbSuccess.Name = "pbSuccess";
            this.pbSuccess.Size = new System.Drawing.Size(24, 26);
            this.pbSuccess.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbSuccess.TabIndex = 1;
            this.pbSuccess.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(663, 32);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Importing attribute translations";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ProgressControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlMain);
            this.Controls.Add(this.pnlLeft);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "ProgressControl";
            this.Size = new System.Drawing.Size(752, 82);
            this.pnlLeft.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbProgress)).EndInit();
            this.pnlMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbTotal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbError)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbSuccess)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlLeft;
        private System.Windows.Forms.Panel pnlMain;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.PictureBox pbSuccess;
        private System.Windows.Forms.PictureBox pbError;
        private System.Windows.Forms.PictureBox pbTotal;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.Label lblSuccess;
        private System.Windows.Forms.Label lblCount;
        private System.Windows.Forms.PictureBox pbProgress;
    }
}
