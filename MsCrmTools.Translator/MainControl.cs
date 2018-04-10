using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk.Metadata;
using MsCrmTools.Translator.AppCode;
using MsCrmTools.Translator.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using CrmExceptionHelper = XrmToolBox.CrmExceptionHelper;

namespace MsCrmTools.Translator
{
    public partial class MainControl : PluginControlBase, IGitHubPlugin, IHelpPlugin
    {
        #region variables

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
                        ExportDescriptions = rdbBoth.Checked || rdbDescOnly.Checked
                    };

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
                                MessageBox.Show(this, errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            numberOferrors = 0;
            pbOverall.ForeColor = SystemColors.Highlight;
            pbItem.ForeColor = SystemColors.Highlight;
            SetState(true);

            gbProgress.Visible = true;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "",
                AsyncArgument = txtFilePath.Text,
                Work = (bw, e) =>
                {
                    var engine = new Engine();
                    engine.OnLog += (sender, evt) =>
                    {
                        if (!evt.Success)
                        {
                            numberOferrors++;
                            LogError("{0}\t{1}", evt.SheetName, evt.Message);
                        }
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
                    //gbProgress.Visible = false;
                },
                ProgressChanged = e =>
                {
                    var pInfo = e.UserState as ProgressInfo;
                    //MessageBox.Show($"{pInfo?.Overall}/{pInfo?.Item}/{pInfo?.Message}");
                    if (pInfo != null)
                    {
                        if (pInfo.Item == 1)
                        {
                            pbItem.ForeColor = SystemColors.Highlight;
                        }

                        if (numberOferrors > 0)
                        {
                            pbOverall.ForeColor = Color.Red;
                            pbItem.ForeColor = Color.Red;
                            pbOverall.Invalidate();
                            pbItem.Invalidate();

                            llOpenLog.Visible = true;
                            lblErrors.Text = string.Format(lblErrors.Tag.ToString(), numberOferrors);
                        }

                        if (pInfo.Overall != 0)
                        {
                            pbOverall.Value = pInfo.Overall > pbOverall.Maximum ? pbOverall.Maximum : pInfo.Overall;
                        }

                        if (pInfo.Item != 0)
                        {
                            pbItem.Value = pInfo.Item > pbItem.Maximum ? pbItem.Maximum : pInfo.Item;
                        }

                        if (pInfo.Message != null)
                        {
                            lblProgress.Text = pInfo.Message;
                        }
                    }
                    else
                    {
                        SetWorkingMessage(e.UserState.ToString());
                    }
                }
            });
        }

        private void LoadEntities(bool allEntities)
        {
            Guid solutionId = Guid.Empty;

            if (!allEntities)
            {
                var sPicker = new SolutionPicker(Service);
                if (sPicker.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                solutionId = sPicker.SelectedSolution.First().Id;
            }

            lvEntities.Items.Clear();

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading entities...",
                Work = (bw, e) =>
                {
                    List<EntityMetadata> entities = MetadataHelper.RetrieveEntities(Service, solutionId);
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
                        foreach (EntityMetadata emd in (List<EntityMetadata>)e.Result)
                        {
                            var item = new ListViewItem { Text = emd.DisplayName.UserLocalizedLabel.Label, Tag = emd };
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

        private void SetState(bool isRunning)
        {
            tabPage1.Enabled = !isRunning;
            txtFilePath.Enabled = !isRunning;
            btnBrowseImportFile.Enabled = !isRunning;
            tsbExportTranslations.Enabled = !isRunning;
            tsbImportTranslations.Enabled = !isRunning;
            tsddbLoadEntities.Enabled = !isRunning;
        }

        private void TsbLoadEntitiesClick(object sender, EventArgs e)
        {
            ExecuteMethod(LoadEntities, true);
        }

        private void llOpenLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenLogFile();
        }

        public string RepositoryName => "MscrmTools.Translator";
        public string UserName => "MscrmTools";
        public string HelpUrl => "https://github.com/MscrmTools/MsCrmTools.Translator/wiki";

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