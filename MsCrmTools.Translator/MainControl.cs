using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using MsCrmTools.Translator.AppCode;
using MsCrmTools.Translator.Controls;
using MsCrmTools.Translator.Forms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using CrmExceptionHelper = XrmToolBox.CrmExceptionHelper;

namespace MsCrmTools.Translator
{
    public partial class MainControl : PluginControlBase, IGitHubPlugin, IHelpPlugin
    {
        #region variables

        private Guid _solutionId = Guid.Empty;
        private List<int> lcids;
        private int numberOferrors = 0;

        #endregion variables

        #region Constructor

        /// <summary>
        /// Initializes a new instance of class <see cref="MainControl"/>
        /// </summary>
        public MainControl()
        {
            InitializeComponent();
        }

        #endregion Constructor

        #region Methods

        private void TsbCloseClick(object sender, EventArgs e)
        {
            CloseTool();
        }

        #endregion Methods

        private ProgressControl currentControl;
        public string HelpUrl => "https://github.com/MscrmTools/MsCrmTools.Translator/wiki";

        public string RepositoryName => "MscrmTools.Translator";

        public string UserName => "MscrmTools";

        private void BtnBrowseImportFileClick(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select translation file",
                Filter = "Excel Workbook|*.xlsx"
            };

            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                txtFilePath.Text = ofd.FileName;
            }
        }

        private void BtnCheckAllClick(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lvEntities.Items)
                item.Checked = true;
        }

        private void BtnClearAllClick(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lvEntities.Items)
                item.Checked = false;
        }

        private void BtnExportTranslationsClick(object sender, EventArgs e)
        {
            if (lvEntities.CheckedItems.Count > 0 || chkExportGlobalOptSet.Checked || chkExportSiteMap.Checked ||
                chkExportDashboards.Checked)
            {
                var entities =
                    (from ListViewItem item in lvEntities.CheckedItems select ((EntityMetadata)item.Tag).LogicalName)
                        .ToList();

                var sfd = new SaveFileDialog { Filter = "Excel workbook|*.xlsx", Title = "Select file destination" };
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    var settings = new ExportSettings
                    {
                        ExportAttributes = chkExportAttributes.Checked,
                        ExportBooleans = chkExportBooleans.Checked,
                        ExportEntities = chkExportEntity.Checked,
                        ExportForms = chkExportForms.Checked,
                        ExportFormFields = chkExportFormsFields.Checked,
                        ExportFormSections = chkExportFormsSections.Checked,
                        ExportFormTabs = chkExportFormsTabs.Checked,
                        ExportGlobalOptionSet = chkExportGlobalOptSet.Checked,
                        ExportOptionSet = chkExportPicklists.Checked,
                        ExportViews = chkExportViews.Checked,
                        ExportCharts = chkExportCharts.Checked,
                        ExportCustomizedRelationships = chkExportCustomizedRelationships.Checked,
                        ExportSiteMap = chkExportSiteMap.Checked,
                        ExportDashboards = chkExportDashboards.Checked,
                        FilePath = sfd.FileName,
                        Entities = entities,
                        ExportNames = rdbBoth.Checked || rdbNameOnly.Checked,
                        ExportDescriptions = rdbBoth.Checked || rdbDescOnly.Checked,
                        SolutionId = _solutionId,
                        LanguageToExport = rdbExportSpecificLanguage.Checked ? ((Language)ccbLanguageToExport.SelectedItem).Lcid : -1
                    };

                    pnlNewProgress.Controls.Clear();

                    SetState(true);

                    WorkAsync(new WorkAsyncInfo
                    {
                        Message = "Exporting Translations...",
                        AsyncArgument = settings,
                        Work = (bw, evt) =>
                        {
                            var engine = new Engine();
                            engine.Export((ExportSettings)evt.Argument, Service, bw);
                        },
                        PostWorkCallBack = evt =>
                        {
                            SetState(false);

                            if (evt.Error != null)
                            {
                                string errorMessage = CrmExceptionHelper.GetErrorMessage(evt.Error, true);
                                MessageBox.Show(this, errorMessage, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                if (DialogResult.Yes == MessageBox.Show(this, @"Do you want to open generated document?", @"Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                                {
                                    Process.Start(settings.FilePath);
                                }
                            }
                        },
                        ProgressChanged = evt => { SetWorkingMessage(evt.UserState.ToString()); }
                    });
                }
            }
        }

        private void BtnImportTranslationsClick(object sender, EventArgs e)
        {
            if (txtFilePath.Text.Length == 0)
            {
                MessageBox.Show(this, "Please select a file to import", "Warning", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                if (tabControl1.SelectedIndex != 1)
                    tabControl1.SelectedIndex = 1;

                return;
            }

            ExecuteMethod(ImportTranslations);
        }

        private void ImportTranslations()
        {
            var lm = new LogManager(GetType());
            if (File.Exists(lm.FilePath))
            {
                if (MessageBox.Show(this, @"A log file already exists. Would you like to create a new log file?", @"Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    File.Copy(lm.FilePath, $"{lm.FilePath.Substring(0, lm.FilePath.Length - 4)}-{DateTime.Now:yyyyMMdd_HHmmss}.txt", true);
                    File.Delete(lm.FilePath);
                }
            }

            numberOferrors = 0;
            pbOverall.ForeColor = SystemColors.Highlight;
            SetState(true);

            WorkAsync(new WorkAsyncInfo
            {
                Message = "",
                AsyncArgument = txtFilePath.Text,
                Work = (bw, e) =>
                {
                    var engine = new Engine();
                    engine.OnLog += (sender, evt) =>
                    {
                        if (evt.Type == LogType.Error)
                        {
                            LogError(evt.Message);
                        }
                        else if (evt.Type == LogType.Warning)
                        {
                            LogWarning(evt.Message);
                        }
                        else
                        {
                            LogInfo(evt.Message);
                        }
                    };
                    engine.OnProgress += (sender, evt) =>
                    {
                        Invoke(new Action(() =>
                        {
                            if (string.IsNullOrEmpty(evt.SheetName))
                            {
                                currentControl?.End(currentControl.Error == 0);
                                return;
                            }

                            if (currentControl == null || currentControl.SheetName != evt.SheetName)
                            {
                                currentControl?.End(currentControl.Error == 0);
                                currentControl = new ProgressControl(evt.SheetName);
                                currentControl.Dock = DockStyle.Top;

                                pnlNewProgress.Controls.Add(currentControl);
                                pnlNewProgress.Controls.SetChildIndex(currentControl, 0);
                                pnlNewProgress.ScrollControlIntoView(currentControl);
                            }

                            currentControl.Count = evt.TotalItems;
                            currentControl.Error = evt.FailureCount;
                            currentControl.Success = evt.SuccessCount;
                        }));
                    };

                    engine.Import(e.Argument.ToString(), Service, bw);
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        string errorMessage = CrmExceptionHelper.GetErrorMessage(e.Error, true);
                        MessageBox.Show(this, errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    SetState(false);

                    lblProgress.Text = "";

                    if (File.Exists(lm.FilePath))
                    {
                        llOpenLog.Visible = true;
                    }
                },
                ProgressChanged = e =>
                {
                    if (e.UserState is ProgressInfo pInfo)
                    {
                        lblProgress.Text = pInfo.Message;
                        if (numberOferrors > 0)
                        {
                            pbOverall.ForeColor = Color.Red;
                            pbOverall.Invalidate();

                            llOpenLog.Visible = true;
                        }

                        if (pInfo.Overall != 0)
                        {
                            pbOverall.Value = pInfo.Overall > pbOverall.Maximum ? pbOverall.Maximum : pInfo.Overall;
                        }
                    }
                    else
                    {
                        SetWorkingMessage(e.UserState.ToString());
                    }
                }
            });
        }

        private void llGlobalSelector_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            bool newStatus;
            if (llGlobalSelector.Text == "Clear all")
            {
                newStatus = false;
                llGlobalSelector.Text = "Select all";
            }
            else
            {
                newStatus = true;
                llGlobalSelector.Text = "Clear all";
            }

            foreach (var ctrl in gbGlobalOptions.Controls)
            {
                var cb = ctrl as CheckBox;
                if (cb != null)
                {
                    cb.Checked = newStatus;
                }
            }
        }

        private void llOpenLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenLogFile();
        }

        private void llRelatedSelector_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            bool newStatus;
            if (llRelatedSelector.Text == "Clear all")
            {
                newStatus = false;
                llRelatedSelector.Text = "Select all";
            }
            else
            {
                newStatus = true;
                llRelatedSelector.Text = "Clear all";
            }

            foreach (var ctrl in gbEntitiesOptions.Controls)
            {
                var cb = ctrl as CheckBox;
                if (cb != null)
                {
                    cb.Checked = newStatus;
                }
            }
        }

        private void LoadEntities(bool allEntities)
        {
            _solutionId = Guid.Empty;
            pnlSelectedSolution.Visible = false;

            if (!allEntities)
            {
                var sPicker = new SolutionPicker(Service);
                if (sPicker.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _solutionId = sPicker.SelectedSolution.First().Id;

                lblSelectedSolution.Text = string.Format(lblSelectedSolution.Tag.ToString(),
                    string.Join(", ",
                        sPicker.SelectedSolution.Select(s => s.GetAttributeValue<string>("friendlyname"))));
                pnlSelectedSolution.Visible = true;
            }

            lvEntities.Items.Clear();

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading entities...",
                Work = (bw, e) =>
                {
                    var lcidRequest = new RetrieveProvisionedLanguagesRequest();
                    var lcidResponse = (RetrieveProvisionedLanguagesResponse)Service.Execute(lcidRequest);
                    lcids = lcidResponse.RetrieveProvisionedLanguages.ToList();

                    List<EntityMetadata> entities = MetadataHelper.RetrieveEntities(Service, _solutionId);
                    e.Result = entities;
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        string errorMessage = CrmExceptionHelper.GetErrorMessage(e.Error, true);
                        MessageBox.Show(this, errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        ccbLanguageToExport.Items.Clear();
                        foreach (var lcid in lcids)
                        {
                            ccbLanguageToExport.Items.Add(new Language(lcid, new CultureInfo(lcid).EnglishName));
                        }

                        ccbLanguageToExport.SelectedIndex = 0;

                        foreach (EntityMetadata emd in (List<EntityMetadata>)e.Result)
                        {
                            var item = new ListViewItem { Text = emd.DisplayName?.UserLocalizedLabel?.Label ?? "N/A", Tag = emd };
                            item.SubItems.Add(emd.LogicalName);
                            lvEntities.Items.Add(item);
                        }
                    }
                }
            });
        }

        private void LvEntitiesColumnClick(object sender, ColumnClickEventArgs e)
        {
            lvEntities.Sorting = lvEntities.Sorting == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            lvEntities.ListViewItemSorter = new ListViewItemComparer(e.Column, lvEntities.Sorting);
        }

        private void rdbExportSpecificLanguage_CheckedChanged(object sender, EventArgs e)
        {
            ccbLanguageToExport.Enabled = rdbExportSpecificLanguage.Checked;
        }

        private void SetState(bool isRunning)
        {
            tabPage1.Enabled = !isRunning;
            txtFilePath.Enabled = !isRunning;
            btnBrowseImportFile.Enabled = !isRunning;
            tsbExportTranslations.Enabled = !isRunning;
            tsbImportTranslations.Enabled = !isRunning;
            tsddbLoadEntities.Enabled = !isRunning;
            pbOverall.Visible = isRunning;
            lblProgress.Visible = isRunning;
        }

        private void tsddbLoadEntities_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem == tsmiAllEntities)
            {
                ExecuteMethod(LoadEntities, true);
            }
            else if (e.ClickedItem == tsmiEntitiesFromASolution)
            {
                ExecuteMethod(LoadEntities, false);
            }
        }
    }
}