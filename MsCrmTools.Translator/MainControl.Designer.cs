﻿using System.Drawing;

namespace MsCrmTools.Translator
{
    partial class MainControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainControl));
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsddbLoadEntities = new System.Windows.Forms.ToolStripDropDownButton();
            this.tsmiAllEntities = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiEntitiesFromASolution = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbExportTranslations = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbImportTranslations = new System.Windows.Forms.ToolStripButton();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.gbLangages = new System.Windows.Forms.GroupBox();
            this.ccbLanguageToExport = new System.Windows.Forms.ComboBox();
            this.rdbExportSpecificLanguage = new System.Windows.Forms.RadioButton();
            this.rdbExportAllLanguages = new System.Windows.Forms.RadioButton();
            this.gbLabelExportOptions = new System.Windows.Forms.GroupBox();
            this.rdbDescOnly = new System.Windows.Forms.RadioButton();
            this.rdbNameOnly = new System.Windows.Forms.RadioButton();
            this.rdbBoth = new System.Windows.Forms.RadioButton();
            this.lblExportLabel = new System.Windows.Forms.Label();
            this.btnCheckAll = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.gbEntitiesOptions = new System.Windows.Forms.GroupBox();
            this.llRelatedSelector = new System.Windows.Forms.LinkLabel();
            this.chkExportCharts = new System.Windows.Forms.CheckBox();
            this.chkExportCustomizedRelationships = new System.Windows.Forms.CheckBox();
            this.chkExportEntity = new System.Windows.Forms.CheckBox();
            this.chkExportAttributes = new System.Windows.Forms.CheckBox();
            this.chkExportFormsFields = new System.Windows.Forms.CheckBox();
            this.chkExportPicklists = new System.Windows.Forms.CheckBox();
            this.chkExportFormsSections = new System.Windows.Forms.CheckBox();
            this.chkExportBooleans = new System.Windows.Forms.CheckBox();
            this.chkExportFormsTabs = new System.Windows.Forms.CheckBox();
            this.chkExportViews = new System.Windows.Forms.CheckBox();
            this.chkExportForms = new System.Windows.Forms.CheckBox();
            this.gbGlobalOptions = new System.Windows.Forms.GroupBox();
            this.llGlobalSelector = new System.Windows.Forms.LinkLabel();
            this.chkExportDashboards = new System.Windows.Forms.CheckBox();
            this.chkExportSiteMap = new System.Windows.Forms.CheckBox();
            this.chkExportGlobalOptSet = new System.Windows.Forms.CheckBox();
            this.lvEntities = new System.Windows.Forms.ListView();
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.pnlSelectedSolution = new System.Windows.Forms.Panel();
            this.lblSelectedSolution = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.pnlImportMain = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.gbSettings = new System.Windows.Forms.GroupBox();
            this.nudBatchCount = new System.Windows.Forms.NumericUpDown();
            this.lblBatchCount = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.gbNewProgress = new System.Windows.Forms.GroupBox();
            this.pnlNewProgress = new System.Windows.Forms.Panel();
            this.llOpenLog = new System.Windows.Forms.LinkLabel();
            this.lblProgress = new System.Windows.Forms.Label();
            this.pbOverall = new MsCrmTools.Translator.ColoredProgressBar();
            this.pnlImportBrowse = new System.Windows.Forms.Panel();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowseImportFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.toolStripMenu.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.gbLangages.SuspendLayout();
            this.gbLabelExportOptions.SuspendLayout();
            this.gbEntitiesOptions.SuspendLayout();
            this.gbGlobalOptions.SuspendLayout();
            this.pnlSelectedSolution.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.pnlImportMain.SuspendLayout();
            this.panel3.SuspendLayout();
            this.gbSettings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudBatchCount)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.gbNewProgress.SuspendLayout();
            this.pnlImportBrowse.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripMenu
            // 
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbClose,
            this.toolStripSeparator1,
            this.tsddbLoadEntities,
            this.toolStripSeparator2,
            this.tsbExportTranslations,
            this.toolStripSeparator3,
            this.tsbImportTranslations});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(2049, 57);
            this.toolStripMenu.TabIndex = 2;
            this.toolStripMenu.Text = "tsMain";
            // 
            // tsbClose
            // 
            this.tsbClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsbClose.Image = ((System.Drawing.Image)(resources.GetObject("tsbClose.Image")));
            this.tsbClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Size = new System.Drawing.Size(34, 52);
            this.tsbClose.Text = "Close this tool";
            this.tsbClose.Click += new System.EventHandler(this.TsbCloseClick);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 34);
            // 
            // tsddbLoadEntities
            // 
            this.tsddbLoadEntities.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiAllEntities,
            this.tsmiEntitiesFromASolution});
            this.tsddbLoadEntities.Image = global::MsCrmTools.Translator.Properties.Resources.Dynamics16;
            this.tsddbLoadEntities.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsddbLoadEntities.Name = "tsddbLoadEntities";
            this.tsddbLoadEntities.Size = new System.Drawing.Size(154, 29);
            this.tsddbLoadEntities.Text = "Load Entities";
            this.tsddbLoadEntities.DropDownItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.tsddbLoadEntities_DropDownItemClicked);
            // 
            // tsmiAllEntities
            // 
            this.tsmiAllEntities.Name = "tsmiAllEntities";
            this.tsmiAllEntities.Size = new System.Drawing.Size(239, 34);
            this.tsmiAllEntities.Text = "All entities";
            // 
            // tsmiEntitiesFromASolution
            // 
            this.tsmiEntitiesFromASolution.Name = "tsmiEntitiesFromASolution";
            this.tsmiEntitiesFromASolution.Size = new System.Drawing.Size(239, 34);
            this.tsmiEntitiesFromASolution.Text = "From a solution";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 34);
            // 
            // tsbExportTranslations
            // 
            this.tsbExportTranslations.Image = ((System.Drawing.Image)(resources.GetObject("tsbExportTranslations.Image")));
            this.tsbExportTranslations.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbExportTranslations.Name = "tsbExportTranslations";
            this.tsbExportTranslations.Size = new System.Drawing.Size(187, 29);
            this.tsbExportTranslations.Text = "Export translations";
            this.tsbExportTranslations.Click += new System.EventHandler(this.BtnExportTranslationsClick);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 34);
            // 
            // tsbImportTranslations
            // 
            this.tsbImportTranslations.Image = ((System.Drawing.Image)(resources.GetObject("tsbImportTranslations.Image")));
            this.tsbImportTranslations.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbImportTranslations.Name = "tsbImportTranslations";
            this.tsbImportTranslations.Size = new System.Drawing.Size(191, 29);
            this.tsbImportTranslations.Text = "Import translations";
            this.tsbImportTranslations.Click += new System.EventHandler(this.BtnImportTranslationsClick);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(4, 43);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1358, 875);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.pnlSelectedSolution);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Size = new System.Drawing.Size(1350, 842);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Export translations";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.gbLangages);
            this.groupBox2.Controls.Add(this.gbLabelExportOptions);
            this.groupBox2.Controls.Add(this.btnCheckAll);
            this.groupBox2.Controls.Add(this.btnClearAll);
            this.groupBox2.Controls.Add(this.gbEntitiesOptions);
            this.groupBox2.Controls.Add(this.gbGlobalOptions);
            this.groupBox2.Controls.Add(this.lvEntities);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(4, 52);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Size = new System.Drawing.Size(1342, 785);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Entities options";
            // 
            // gbLangages
            // 
            this.gbLangages.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.gbLangages.Controls.Add(this.ccbLanguageToExport);
            this.gbLangages.Controls.Add(this.rdbExportSpecificLanguage);
            this.gbLangages.Controls.Add(this.rdbExportAllLanguages);
            this.gbLangages.Location = new System.Drawing.Point(994, 29);
            this.gbLangages.Margin = new System.Windows.Forms.Padding(2);
            this.gbLangages.Name = "gbLangages";
            this.gbLangages.Padding = new System.Windows.Forms.Padding(2);
            this.gbLangages.Size = new System.Drawing.Size(338, 151);
            this.gbLangages.TabIndex = 97;
            this.gbLangages.TabStop = false;
            this.gbLangages.Text = "Langages";
            // 
            // ccbLanguageToExport
            // 
            this.ccbLanguageToExport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ccbLanguageToExport.Enabled = false;
            this.ccbLanguageToExport.FormattingEnabled = true;
            this.ccbLanguageToExport.Location = new System.Drawing.Point(122, 63);
            this.ccbLanguageToExport.Margin = new System.Windows.Forms.Padding(2);
            this.ccbLanguageToExport.Name = "ccbLanguageToExport";
            this.ccbLanguageToExport.Size = new System.Drawing.Size(206, 28);
            this.ccbLanguageToExport.TabIndex = 2;
            // 
            // rdbExportSpecificLanguage
            // 
            this.rdbExportSpecificLanguage.AutoSize = true;
            this.rdbExportSpecificLanguage.Location = new System.Drawing.Point(4, 63);
            this.rdbExportSpecificLanguage.Margin = new System.Windows.Forms.Padding(2);
            this.rdbExportSpecificLanguage.Name = "rdbExportSpecificLanguage";
            this.rdbExportSpecificLanguage.Size = new System.Drawing.Size(112, 24);
            this.rdbExportSpecificLanguage.TabIndex = 1;
            this.rdbExportSpecificLanguage.Text = "Export only";
            this.rdbExportSpecificLanguage.UseVisualStyleBackColor = true;
            this.rdbExportSpecificLanguage.CheckedChanged += new System.EventHandler(this.rdbExportSpecificLanguage_CheckedChanged);
            // 
            // rdbExportAllLanguages
            // 
            this.rdbExportAllLanguages.AutoSize = true;
            this.rdbExportAllLanguages.Checked = true;
            this.rdbExportAllLanguages.Location = new System.Drawing.Point(4, 29);
            this.rdbExportAllLanguages.Margin = new System.Windows.Forms.Padding(2);
            this.rdbExportAllLanguages.Name = "rdbExportAllLanguages";
            this.rdbExportAllLanguages.Size = new System.Drawing.Size(99, 24);
            this.rdbExportAllLanguages.TabIndex = 0;
            this.rdbExportAllLanguages.TabStop = true;
            this.rdbExportAllLanguages.Text = "Export all";
            this.rdbExportAllLanguages.UseVisualStyleBackColor = true;
            // 
            // gbLabelExportOptions
            // 
            this.gbLabelExportOptions.Controls.Add(this.rdbDescOnly);
            this.gbLabelExportOptions.Controls.Add(this.rdbNameOnly);
            this.gbLabelExportOptions.Controls.Add(this.rdbBoth);
            this.gbLabelExportOptions.Controls.Add(this.lblExportLabel);
            this.gbLabelExportOptions.Location = new System.Drawing.Point(490, 691);
            this.gbLabelExportOptions.Name = "gbLabelExportOptions";
            this.gbLabelExportOptions.Size = new System.Drawing.Size(832, 100);
            this.gbLabelExportOptions.TabIndex = 96;
            this.gbLabelExportOptions.TabStop = false;
            this.gbLabelExportOptions.Text = "Label options";
            // 
            // rdbDescOnly
            // 
            this.rdbDescOnly.AutoSize = true;
            this.rdbDescOnly.Location = new System.Drawing.Point(433, 22);
            this.rdbDescOnly.Name = "rdbDescOnly";
            this.rdbDescOnly.Size = new System.Drawing.Size(157, 24);
            this.rdbDescOnly.TabIndex = 3;
            this.rdbDescOnly.TabStop = true;
            this.rdbDescOnly.Text = "Only Descriptions";
            this.rdbDescOnly.UseVisualStyleBackColor = true;
            // 
            // rdbNameOnly
            // 
            this.rdbNameOnly.AutoSize = true;
            this.rdbNameOnly.Location = new System.Drawing.Point(308, 22);
            this.rdbNameOnly.Name = "rdbNameOnly";
            this.rdbNameOnly.Size = new System.Drawing.Size(119, 24);
            this.rdbNameOnly.TabIndex = 2;
            this.rdbNameOnly.TabStop = true;
            this.rdbNameOnly.Text = "Only Names";
            this.rdbNameOnly.UseVisualStyleBackColor = true;
            // 
            // rdbBoth
            // 
            this.rdbBoth.AutoSize = true;
            this.rdbBoth.Checked = true;
            this.rdbBoth.Location = new System.Drawing.Point(234, 22);
            this.rdbBoth.Name = "rdbBoth";
            this.rdbBoth.Size = new System.Drawing.Size(68, 24);
            this.rdbBoth.TabIndex = 1;
            this.rdbBoth.TabStop = true;
            this.rdbBoth.Text = "Both";
            this.rdbBoth.UseVisualStyleBackColor = true;
            // 
            // lblExportLabel
            // 
            this.lblExportLabel.AutoSize = true;
            this.lblExportLabel.Location = new System.Drawing.Point(6, 24);
            this.lblExportLabel.Name = "lblExportLabel";
            this.lblExportLabel.Size = new System.Drawing.Size(224, 20);
            this.lblExportLabel.TabIndex = 0;
            this.lblExportLabel.Text = "Export Name and Description :";
            // 
            // btnCheckAll
            // 
            this.btnCheckAll.Location = new System.Drawing.Point(248, 29);
            this.btnCheckAll.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCheckAll.Name = "btnCheckAll";
            this.btnCheckAll.Size = new System.Drawing.Size(112, 35);
            this.btnCheckAll.TabIndex = 95;
            this.btnCheckAll.Text = "Check all";
            this.btnCheckAll.UseVisualStyleBackColor = true;
            this.btnCheckAll.Click += new System.EventHandler(this.BtnCheckAllClick);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(369, 29);
            this.btnClearAll.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(112, 35);
            this.btnClearAll.TabIndex = 94;
            this.btnClearAll.Text = "Clear all";
            this.btnClearAll.UseVisualStyleBackColor = true;
            this.btnClearAll.Click += new System.EventHandler(this.BtnClearAllClick);
            // 
            // gbEntitiesOptions
            // 
            this.gbEntitiesOptions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbEntitiesOptions.Controls.Add(this.llRelatedSelector);
            this.gbEntitiesOptions.Controls.Add(this.chkExportCharts);
            this.gbEntitiesOptions.Controls.Add(this.chkExportCustomizedRelationships);
            this.gbEntitiesOptions.Controls.Add(this.chkExportEntity);
            this.gbEntitiesOptions.Controls.Add(this.chkExportAttributes);
            this.gbEntitiesOptions.Controls.Add(this.chkExportFormsFields);
            this.gbEntitiesOptions.Controls.Add(this.chkExportPicklists);
            this.gbEntitiesOptions.Controls.Add(this.chkExportFormsSections);
            this.gbEntitiesOptions.Controls.Add(this.chkExportBooleans);
            this.gbEntitiesOptions.Controls.Add(this.chkExportFormsTabs);
            this.gbEntitiesOptions.Controls.Add(this.chkExportViews);
            this.gbEntitiesOptions.Controls.Add(this.chkExportForms);
            this.gbEntitiesOptions.Location = new System.Drawing.Point(490, 194);
            this.gbEntitiesOptions.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.gbEntitiesOptions.MinimumSize = new System.Drawing.Size(376, 382);
            this.gbEntitiesOptions.Name = "gbEntitiesOptions";
            this.gbEntitiesOptions.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.gbEntitiesOptions.Size = new System.Drawing.Size(841, 483);
            this.gbEntitiesOptions.TabIndex = 93;
            this.gbEntitiesOptions.TabStop = false;
            this.gbEntitiesOptions.Text = "Entity related options";
            // 
            // llRelatedSelector
            // 
            this.llRelatedSelector.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.llRelatedSelector.AutoSize = true;
            this.llRelatedSelector.Location = new System.Drawing.Point(763, 25);
            this.llRelatedSelector.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llRelatedSelector.Name = "llRelatedSelector";
            this.llRelatedSelector.Size = new System.Drawing.Size(65, 20);
            this.llRelatedSelector.TabIndex = 95;
            this.llRelatedSelector.TabStop = true;
            this.llRelatedSelector.Text = "Clear all";
            this.llRelatedSelector.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llRelatedSelector_LinkClicked);
            // 
            // chkExportCharts
            // 
            this.chkExportCharts.AutoSize = true;
            this.chkExportCharts.Checked = true;
            this.chkExportCharts.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportCharts.Location = new System.Drawing.Point(9, 206);
            this.chkExportCharts.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportCharts.Name = "chkExportCharts";
            this.chkExportCharts.Size = new System.Drawing.Size(214, 24);
            this.chkExportCharts.TabIndex = 91;
            this.chkExportCharts.Text = "Export Charts Translation";
            this.chkExportCharts.UseVisualStyleBackColor = true;
            // 
            // chkExportCustomizedRelationships
            // 
            this.chkExportCustomizedRelationships.AutoSize = true;
            this.chkExportCustomizedRelationships.Checked = true;
            this.chkExportCustomizedRelationships.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportCustomizedRelationships.Location = new System.Drawing.Point(9, 382);
            this.chkExportCustomizedRelationships.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportCustomizedRelationships.Name = "chkExportCustomizedRelationships";
            this.chkExportCustomizedRelationships.Size = new System.Drawing.Size(383, 24);
            this.chkExportCustomizedRelationships.TabIndex = 90;
            this.chkExportCustomizedRelationships.Text = "Export Relationships that are using custom labels";
            this.chkExportCustomizedRelationships.UseVisualStyleBackColor = true;
            // 
            // chkExportEntity
            // 
            this.chkExportEntity.AutoSize = true;
            this.chkExportEntity.Checked = true;
            this.chkExportEntity.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportEntity.Location = new System.Drawing.Point(9, 29);
            this.chkExportEntity.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportEntity.Name = "chkExportEntity";
            this.chkExportEntity.Size = new System.Drawing.Size(207, 24);
            this.chkExportEntity.TabIndex = 81;
            this.chkExportEntity.Text = "Export Entity Translation";
            this.chkExportEntity.UseVisualStyleBackColor = true;
            // 
            // chkExportAttributes
            // 
            this.chkExportAttributes.AutoSize = true;
            this.chkExportAttributes.Checked = true;
            this.chkExportAttributes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportAttributes.Location = new System.Drawing.Point(9, 65);
            this.chkExportAttributes.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportAttributes.Name = "chkExportAttributes";
            this.chkExportAttributes.Size = new System.Drawing.Size(236, 24);
            this.chkExportAttributes.TabIndex = 82;
            this.chkExportAttributes.Text = "Export Attributes Translation";
            this.chkExportAttributes.UseVisualStyleBackColor = true;
            // 
            // chkExportFormsFields
            // 
            this.chkExportFormsFields.AutoSize = true;
            this.chkExportFormsFields.Checked = true;
            this.chkExportFormsFields.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportFormsFields.Location = new System.Drawing.Point(9, 348);
            this.chkExportFormsFields.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportFormsFields.Name = "chkExportFormsFields";
            this.chkExportFormsFields.Size = new System.Drawing.Size(258, 24);
            this.chkExportFormsFields.TabIndex = 89;
            this.chkExportFormsFields.Text = "Export Forms Fields Translation";
            this.chkExportFormsFields.UseVisualStyleBackColor = true;
            // 
            // chkExportPicklists
            // 
            this.chkExportPicklists.AutoSize = true;
            this.chkExportPicklists.Checked = true;
            this.chkExportPicklists.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportPicklists.Location = new System.Drawing.Point(9, 100);
            this.chkExportPicklists.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportPicklists.Name = "chkExportPicklists";
            this.chkExportPicklists.Size = new System.Drawing.Size(325, 24);
            this.chkExportPicklists.TabIndex = 83;
            this.chkExportPicklists.Text = "Export Picklists Option Labels Translation";
            this.chkExportPicklists.UseVisualStyleBackColor = true;
            // 
            // chkExportFormsSections
            // 
            this.chkExportFormsSections.AutoSize = true;
            this.chkExportFormsSections.Checked = true;
            this.chkExportFormsSections.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportFormsSections.Location = new System.Drawing.Point(9, 312);
            this.chkExportFormsSections.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportFormsSections.Name = "chkExportFormsSections";
            this.chkExportFormsSections.Size = new System.Drawing.Size(278, 24);
            this.chkExportFormsSections.TabIndex = 88;
            this.chkExportFormsSections.Text = "Export Forms Sections Translation";
            this.chkExportFormsSections.UseVisualStyleBackColor = true;
            // 
            // chkExportBooleans
            // 
            this.chkExportBooleans.AutoSize = true;
            this.chkExportBooleans.Checked = true;
            this.chkExportBooleans.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportBooleans.Location = new System.Drawing.Point(9, 135);
            this.chkExportBooleans.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportBooleans.Name = "chkExportBooleans";
            this.chkExportBooleans.Size = new System.Drawing.Size(336, 24);
            this.chkExportBooleans.TabIndex = 84;
            this.chkExportBooleans.Text = "Export Booleans Option Labels Translation";
            this.chkExportBooleans.UseVisualStyleBackColor = true;
            // 
            // chkExportFormsTabs
            // 
            this.chkExportFormsTabs.AutoSize = true;
            this.chkExportFormsTabs.Checked = true;
            this.chkExportFormsTabs.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportFormsTabs.Location = new System.Drawing.Point(9, 277);
            this.chkExportFormsTabs.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportFormsTabs.Name = "chkExportFormsTabs";
            this.chkExportFormsTabs.Size = new System.Drawing.Size(251, 24);
            this.chkExportFormsTabs.TabIndex = 87;
            this.chkExportFormsTabs.Text = "Export Forms Tabs Translation";
            this.chkExportFormsTabs.UseVisualStyleBackColor = true;
            // 
            // chkExportViews
            // 
            this.chkExportViews.AutoSize = true;
            this.chkExportViews.Checked = true;
            this.chkExportViews.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportViews.Location = new System.Drawing.Point(9, 171);
            this.chkExportViews.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportViews.Name = "chkExportViews";
            this.chkExportViews.Size = new System.Drawing.Size(209, 24);
            this.chkExportViews.TabIndex = 85;
            this.chkExportViews.Text = "Export Views Translation";
            this.chkExportViews.UseVisualStyleBackColor = true;
            // 
            // chkExportForms
            // 
            this.chkExportForms.AutoSize = true;
            this.chkExportForms.Checked = true;
            this.chkExportForms.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportForms.Location = new System.Drawing.Point(9, 242);
            this.chkExportForms.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportForms.Name = "chkExportForms";
            this.chkExportForms.Size = new System.Drawing.Size(212, 24);
            this.chkExportForms.TabIndex = 86;
            this.chkExportForms.Text = "Export Forms Translation";
            this.chkExportForms.UseVisualStyleBackColor = true;
            // 
            // gbGlobalOptions
            // 
            this.gbGlobalOptions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbGlobalOptions.Controls.Add(this.llGlobalSelector);
            this.gbGlobalOptions.Controls.Add(this.chkExportDashboards);
            this.gbGlobalOptions.Controls.Add(this.chkExportSiteMap);
            this.gbGlobalOptions.Controls.Add(this.chkExportGlobalOptSet);
            this.gbGlobalOptions.Location = new System.Drawing.Point(490, 29);
            this.gbGlobalOptions.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.gbGlobalOptions.MinimumSize = new System.Drawing.Size(376, 108);
            this.gbGlobalOptions.Name = "gbGlobalOptions";
            this.gbGlobalOptions.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.gbGlobalOptions.Size = new System.Drawing.Size(498, 151);
            this.gbGlobalOptions.TabIndex = 92;
            this.gbGlobalOptions.TabStop = false;
            this.gbGlobalOptions.Text = "Global Options";
            // 
            // llGlobalSelector
            // 
            this.llGlobalSelector.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.llGlobalSelector.AutoSize = true;
            this.llGlobalSelector.Location = new System.Drawing.Point(419, 15);
            this.llGlobalSelector.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llGlobalSelector.Name = "llGlobalSelector";
            this.llGlobalSelector.Size = new System.Drawing.Size(65, 20);
            this.llGlobalSelector.TabIndex = 94;
            this.llGlobalSelector.TabStop = true;
            this.llGlobalSelector.Text = "Clear all";
            this.llGlobalSelector.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llGlobalSelector_LinkClicked);
            // 
            // chkExportDashboards
            // 
            this.chkExportDashboards.AutoSize = true;
            this.chkExportDashboards.Checked = true;
            this.chkExportDashboards.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportDashboards.Location = new System.Drawing.Point(9, 97);
            this.chkExportDashboards.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportDashboards.Name = "chkExportDashboards";
            this.chkExportDashboards.Size = new System.Drawing.Size(355, 24);
            this.chkExportDashboards.TabIndex = 93;
            this.chkExportDashboards.Text = "Export Dashboards custom labels Translation";
            this.chkExportDashboards.UseVisualStyleBackColor = true;
            // 
            // chkExportSiteMap
            // 
            this.chkExportSiteMap.AutoSize = true;
            this.chkExportSiteMap.Checked = true;
            this.chkExportSiteMap.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportSiteMap.Location = new System.Drawing.Point(8, 63);
            this.chkExportSiteMap.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportSiteMap.Name = "chkExportSiteMap";
            this.chkExportSiteMap.Size = new System.Drawing.Size(327, 24);
            this.chkExportSiteMap.TabIndex = 92;
            this.chkExportSiteMap.Text = "Export SiteMap custom labels Translation";
            this.chkExportSiteMap.UseVisualStyleBackColor = true;
            // 
            // chkExportGlobalOptSet
            // 
            this.chkExportGlobalOptSet.AutoSize = true;
            this.chkExportGlobalOptSet.Checked = true;
            this.chkExportGlobalOptSet.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExportGlobalOptSet.Location = new System.Drawing.Point(9, 29);
            this.chkExportGlobalOptSet.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkExportGlobalOptSet.Name = "chkExportGlobalOptSet";
            this.chkExportGlobalOptSet.Size = new System.Drawing.Size(348, 24);
            this.chkExportGlobalOptSet.TabIndex = 91;
            this.chkExportGlobalOptSet.Text = "Export Global OptionSets Labels Translation";
            this.chkExportGlobalOptSet.UseVisualStyleBackColor = true;
            // 
            // lvEntities
            // 
            this.lvEntities.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.lvEntities.CheckBoxes = true;
            this.lvEntities.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader4,
            this.columnHeader7});
            this.lvEntities.FullRowSelect = true;
            this.lvEntities.HideSelection = false;
            this.lvEntities.Location = new System.Drawing.Point(10, 74);
            this.lvEntities.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.lvEntities.Name = "lvEntities";
            this.lvEntities.Size = new System.Drawing.Size(469, 700);
            this.lvEntities.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.lvEntities.TabIndex = 80;
            this.lvEntities.UseCompatibleStateImageBehavior = false;
            this.lvEntities.View = System.Windows.Forms.View.Details;
            this.lvEntities.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.LvEntitiesColumnClick);
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Display name";
            this.columnHeader4.Width = 157;
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "Schema name";
            this.columnHeader7.Width = 130;
            // 
            // pnlSelectedSolution
            // 
            this.pnlSelectedSolution.Controls.Add(this.lblSelectedSolution);
            this.pnlSelectedSolution.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlSelectedSolution.Location = new System.Drawing.Point(4, 5);
            this.pnlSelectedSolution.Margin = new System.Windows.Forms.Padding(2);
            this.pnlSelectedSolution.Name = "pnlSelectedSolution";
            this.pnlSelectedSolution.Padding = new System.Windows.Forms.Padding(8);
            this.pnlSelectedSolution.Size = new System.Drawing.Size(1342, 47);
            this.pnlSelectedSolution.TabIndex = 7;
            this.pnlSelectedSolution.Visible = false;
            // 
            // lblSelectedSolution
            // 
            this.lblSelectedSolution.BackColor = System.Drawing.SystemColors.Info;
            this.lblSelectedSolution.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSelectedSolution.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblSelectedSolution.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelectedSolution.Location = new System.Drawing.Point(8, 8);
            this.lblSelectedSolution.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblSelectedSolution.Name = "lblSelectedSolution";
            this.lblSelectedSolution.Size = new System.Drawing.Size(1326, 31);
            this.lblSelectedSolution.TabIndex = 0;
            this.lblSelectedSolution.Tag = "Selected solution : {0}. Global OptionSets, SiteMap and Dashboards will be export" +
    "ed for this solution only";
            this.lblSelectedSolution.Text = "Selected solution : {0}";
            this.lblSelectedSolution.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.pnlImportMain);
            this.tabPage2.Controls.Add(this.pnlImportBrowse);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage2.Size = new System.Drawing.Size(1350, 842);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Import translations";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // pnlImportMain
            // 
            this.pnlImportMain.Controls.Add(this.panel3);
            this.pnlImportMain.Controls.Add(this.gbNewProgress);
            this.pnlImportMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlImportMain.Location = new System.Drawing.Point(4, 60);
            this.pnlImportMain.Margin = new System.Windows.Forms.Padding(2);
            this.pnlImportMain.Name = "pnlImportMain";
            this.pnlImportMain.Size = new System.Drawing.Size(1342, 777);
            this.pnlImportMain.TabIndex = 10;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.gbSettings);
            this.panel3.Controls.Add(this.panel2);
            this.panel3.Controls.Add(this.panel1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Margin = new System.Windows.Forms.Padding(2);
            this.panel3.Name = "panel3";
            this.panel3.Padding = new System.Windows.Forms.Padding(8);
            this.panel3.Size = new System.Drawing.Size(914, 777);
            this.panel3.TabIndex = 9;
            // 
            // gbSettings
            // 
            this.gbSettings.Controls.Add(this.nudBatchCount);
            this.gbSettings.Controls.Add(this.lblBatchCount);
            this.gbSettings.Dock = System.Windows.Forms.DockStyle.Top;
            this.gbSettings.Location = new System.Drawing.Point(8, 323);
            this.gbSettings.Name = "gbSettings";
            this.gbSettings.Size = new System.Drawing.Size(898, 100);
            this.gbSettings.TabIndex = 5;
            this.gbSettings.TabStop = false;
            this.gbSettings.Text = "Import Settings";
            // 
            // nudBatchCount
            // 
            this.nudBatchCount.Increment = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.nudBatchCount.Location = new System.Drawing.Point(107, 33);
            this.nudBatchCount.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nudBatchCount.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.nudBatchCount.Name = "nudBatchCount";
            this.nudBatchCount.Size = new System.Drawing.Size(78, 26);
            this.nudBatchCount.TabIndex = 1;
            this.nudBatchCount.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // lblBatchCount
            // 
            this.lblBatchCount.AutoSize = true;
            this.lblBatchCount.Location = new System.Drawing.Point(6, 35);
            this.lblBatchCount.Name = "lblBatchCount";
            this.lblBatchCount.Size = new System.Drawing.Size(95, 20);
            this.lblBatchCount.TabIndex = 0;
            this.lblBatchCount.Text = "Batch count";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.LightYellow;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.pictureBox2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(8, 221);
            this.panel2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(898, 102);
            this.panel2.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(62, 18);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.MinimumSize = new System.Drawing.Size(834, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(836, 82);
            this.label3.TabIndex = 1;
            this.label3.Text = "Prior to using this tool ensure that you have a backup of the current customizati" +
    "ons\r\nIf there is a problem the backup can be re-imported to restore original tra" +
    "nslations";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(4, 18);
            this.pictureBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(48, 49);
            this.pictureBox2.TabIndex = 0;
            this.pictureBox2.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightYellow;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(8, 8);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(898, 213);
            this.panel1.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(62, 18);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.MinimumSize = new System.Drawing.Size(735, 188);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(737, 188);
            this.label2.TabIndex = 1;
            this.label2.Text = resources.GetString("label2.Text");
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(4, 18);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(48, 49);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // gbNewProgress
            // 
            this.gbNewProgress.Controls.Add(this.pnlNewProgress);
            this.gbNewProgress.Controls.Add(this.llOpenLog);
            this.gbNewProgress.Controls.Add(this.lblProgress);
            this.gbNewProgress.Controls.Add(this.pbOverall);
            this.gbNewProgress.Dock = System.Windows.Forms.DockStyle.Right;
            this.gbNewProgress.Location = new System.Drawing.Point(914, 0);
            this.gbNewProgress.Margin = new System.Windows.Forms.Padding(2);
            this.gbNewProgress.Name = "gbNewProgress";
            this.gbNewProgress.Padding = new System.Windows.Forms.Padding(2);
            this.gbNewProgress.Size = new System.Drawing.Size(428, 777);
            this.gbNewProgress.TabIndex = 8;
            this.gbNewProgress.TabStop = false;
            this.gbNewProgress.Text = "Progress";
            // 
            // pnlNewProgress
            // 
            this.pnlNewProgress.AutoScroll = true;
            this.pnlNewProgress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlNewProgress.Location = new System.Drawing.Point(2, 115);
            this.pnlNewProgress.Margin = new System.Windows.Forms.Padding(2);
            this.pnlNewProgress.Name = "pnlNewProgress";
            this.pnlNewProgress.Size = new System.Drawing.Size(424, 660);
            this.pnlNewProgress.TabIndex = 16;
            // 
            // llOpenLog
            // 
            this.llOpenLog.Dock = System.Windows.Forms.DockStyle.Top;
            this.llOpenLog.Location = new System.Drawing.Point(2, 80);
            this.llOpenLog.Name = "llOpenLog";
            this.llOpenLog.Size = new System.Drawing.Size(424, 35);
            this.llOpenLog.TabIndex = 6;
            this.llOpenLog.TabStop = true;
            this.llOpenLog.Text = "Open log file";
            this.llOpenLog.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.llOpenLog.Visible = false;
            this.llOpenLog.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llOpenLog_LinkClicked);
            // 
            // lblProgress
            // 
            this.lblProgress.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblProgress.Location = new System.Drawing.Point(2, 44);
            this.lblProgress.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(424, 36);
            this.lblProgress.TabIndex = 14;
            this.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pbOverall
            // 
            this.pbOverall.Dock = System.Windows.Forms.DockStyle.Top;
            this.pbOverall.Location = new System.Drawing.Point(2, 21);
            this.pbOverall.Name = "pbOverall";
            this.pbOverall.Size = new System.Drawing.Size(424, 23);
            this.pbOverall.TabIndex = 12;
            this.pbOverall.Visible = false;
            // 
            // pnlImportBrowse
            // 
            this.pnlImportBrowse.Controls.Add(this.txtFilePath);
            this.pnlImportBrowse.Controls.Add(this.btnBrowseImportFile);
            this.pnlImportBrowse.Controls.Add(this.label1);
            this.pnlImportBrowse.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlImportBrowse.Location = new System.Drawing.Point(4, 5);
            this.pnlImportBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.pnlImportBrowse.Name = "pnlImportBrowse";
            this.pnlImportBrowse.Padding = new System.Windows.Forms.Padding(8);
            this.pnlImportBrowse.Size = new System.Drawing.Size(1342, 55);
            this.pnlImportBrowse.TabIndex = 9;
            // 
            // txtFilePath
            // 
            this.txtFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFilePath.Location = new System.Drawing.Point(146, 16);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(1070, 26);
            this.txtFilePath.TabIndex = 3;
            // 
            // btnBrowseImportFile
            // 
            this.btnBrowseImportFile.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnBrowseImportFile.Location = new System.Drawing.Point(1222, 8);
            this.btnBrowseImportFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnBrowseImportFile.Name = "btnBrowseImportFile";
            this.btnBrowseImportFile.Size = new System.Drawing.Size(112, 39);
            this.btnBrowseImportFile.TabIndex = 2;
            this.btnBrowseImportFile.Text = "...";
            this.btnBrowseImportFile.UseVisualStyleBackColor = true;
            this.btnBrowseImportFile.Click += new System.EventHandler(this.BtnBrowseImportFileClick);
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Left;
            this.label1.Location = new System.Drawing.Point(8, 8);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 39);
            this.label1.TabIndex = 0;
            this.label1.Text = "Translation file";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // MainControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.toolStripMenu);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "MainControl";
            this.Size = new System.Drawing.Size(1366, 923);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.gbLangages.ResumeLayout(false);
            this.gbLangages.PerformLayout();
            this.gbLabelExportOptions.ResumeLayout(false);
            this.gbLabelExportOptions.PerformLayout();
            this.gbEntitiesOptions.ResumeLayout(false);
            this.gbEntitiesOptions.PerformLayout();
            this.gbGlobalOptions.ResumeLayout(false);
            this.gbGlobalOptions.PerformLayout();
            this.pnlSelectedSolution.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.pnlImportMain.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.gbSettings.ResumeLayout(false);
            this.gbSettings.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudBatchCount)).EndInit();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.gbNewProgress.ResumeLayout(false);
            this.pnlImportBrowse.ResumeLayout(false);
            this.pnlImportBrowse.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ListView lvEntities;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.CheckBox chkExportFormsFields;
        private System.Windows.Forms.CheckBox chkExportFormsSections;
        private System.Windows.Forms.CheckBox chkExportFormsTabs;
        private System.Windows.Forms.CheckBox chkExportForms;
        private System.Windows.Forms.CheckBox chkExportViews;
        private System.Windows.Forms.CheckBox chkExportBooleans;
        private System.Windows.Forms.CheckBox chkExportPicklists;
        private System.Windows.Forms.CheckBox chkExportAttributes;
        private System.Windows.Forms.CheckBox chkExportEntity;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.Button btnBrowseImportFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkExportGlobalOptSet;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton tsbExportTranslations;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripButton tsbImportTranslations;
        private System.Windows.Forms.GroupBox gbEntitiesOptions;
        private System.Windows.Forms.GroupBox gbGlobalOptions;
        private System.Windows.Forms.Button btnCheckAll;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.CheckBox chkExportSiteMap;
        private System.Windows.Forms.CheckBox chkExportCustomizedRelationships;
        private System.Windows.Forms.CheckBox chkExportDashboards;
        private System.Windows.Forms.CheckBox chkExportCharts;
        private System.Windows.Forms.LinkLabel llRelatedSelector;
        private System.Windows.Forms.LinkLabel llGlobalSelector;
        private System.Windows.Forms.GroupBox gbLabelExportOptions;
        private System.Windows.Forms.RadioButton rdbDescOnly;
        private System.Windows.Forms.RadioButton rdbNameOnly;
        private System.Windows.Forms.RadioButton rdbBoth;
        private System.Windows.Forms.Label lblExportLabel;
        private System.Windows.Forms.ToolStripDropDownButton tsddbLoadEntities;
        private System.Windows.Forms.ToolStripMenuItem tsmiAllEntities;
        private System.Windows.Forms.ToolStripMenuItem tsmiEntitiesFromASolution;
        private System.Windows.Forms.GroupBox gbNewProgress;
        private System.Windows.Forms.Panel pnlImportMain;
        private System.Windows.Forms.Panel pnlImportBrowse;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.LinkLabel llOpenLog;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private ColoredProgressBar pbOverall;
        private System.Windows.Forms.Panel pnlSelectedSolution;
        private System.Windows.Forms.Label lblSelectedSolution;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Panel pnlNewProgress;
        private System.Windows.Forms.GroupBox gbLangages;
        private System.Windows.Forms.ComboBox ccbLanguageToExport;
        private System.Windows.Forms.RadioButton rdbExportSpecificLanguage;
        private System.Windows.Forms.RadioButton rdbExportAllLanguages;
        private System.Windows.Forms.GroupBox gbSettings;
        private System.Windows.Forms.NumericUpDown nudBatchCount;
        private System.Windows.Forms.Label lblBatchCount;
    }
}
