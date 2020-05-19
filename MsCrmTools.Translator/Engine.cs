using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using MsCrmTools.Translator.AppCode;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace MsCrmTools.Translator
{
    public class Engine
    {
        public event EventHandler<LogEventArgs> OnLog;

        public event EventHandler<TranslationProgressEventArgs> OnProgress;

        public void Export(ExportSettings settings, IOrganizationService service, BackgroundWorker worker = null)
        {
            List<int> lcids;
            if (settings.LanguageToExport != -1)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Loading environment base language...");
                }
                var baseLanguageCode = service.RetrieveMultiple(new QueryExpression("organization")
                {
                    ColumnSet = new ColumnSet("languagecode")
                }).Entities.First().GetAttributeValue<int>("languagecode");

                lcids = new List<int> { baseLanguageCode, settings.LanguageToExport }.Distinct().ToList();
            }
            else
            {
                // Loading available languages
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Loading provisioned languages...");
                }
                var lcidRequest = new RetrieveProvisionedLanguagesRequest();
                var lcidResponse = (RetrieveProvisionedLanguagesResponse)service.Execute(lcidRequest);
                lcids = lcidResponse.RetrieveProvisionedLanguages.Select(lcid => lcid).ToList();
            }

            // Loading entities
            var emds = new List<EntityMetadata>();

            if (worker != null && worker.WorkerReportsProgress)
            {
                worker.ReportProgress(0, "Loading selected entities...");
            }
            foreach (string entityLogicalName in settings.Entities)
            {
                var filters = EntityFilters.Default;
                if (settings.ExportEntities)
                {
                    filters = filters | EntityFilters.Entity;
                }
                if (settings.ExportCustomizedRelationships)
                {
                    filters = filters | EntityFilters.Relationships;
                }
                if (settings.ExportAttributes || settings.ExportOptionSet || settings.ExportBooleans)
                {
                    filters = filters | EntityFilters.Attributes;
                }

                var request = new RetrieveEntityRequest { LogicalName = entityLogicalName, EntityFilters = filters };
                var response = (RetrieveEntityResponse)service.Execute(request);
                emds.Add(response.EntityMetadata);
            }

            var file = new ExcelPackage();
            file.File = new FileInfo(settings.FilePath);

            if (settings.ExportEntities && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting entities translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Entities");
                var et = new EntityTranslation();
                et.Export(emds, lcids, sheet, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if (settings.ExportAttributes && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting attributes translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Attributes");
                var at = new AttributeTranslation();
                at.Export(emds, lcids, sheet, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if (settings.ExportCustomizedRelationships && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting relationships with custom labels translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Relationships");
                var rt = new RelationshipTranslation();
                rt.Export(emds, lcids, sheet);
                StyleMutator.FontDefaults(sheet);

                var sheetNn = file.Workbook.Worksheets.Add("RelationshipsNN");
                var rtNn = new RelationshipNnTranslation();
                rtNn.Export(emds, lcids, sheetNn);
                StyleMutator.FontDefaults(sheetNn);
            }

            if (settings.ExportGlobalOptionSet)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting global optionsets translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Global OptionSets");
                var ot = new GlobalOptionSetTranslation();
                ot.Export(lcids, sheet, service, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if (settings.ExportOptionSet && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting optionset translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("OptionSets");
                var ot = new OptionSetTranslation();
                ot.Export(emds, lcids, sheet, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if (settings.ExportBooleans && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting booleans translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Booleans");

                var bt = new BooleanTranslation();
                bt.Export(emds, lcids, sheet, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if (settings.ExportViews && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting views translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Views");
                var vt = new ViewTranslation();
                vt.Export(emds, lcids, sheet, service, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if (settings.ExportCharts && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting Charts translations...");
                }

                var sheet = file.Workbook.Worksheets.Add("Charts");
                var vt = new VisualizationTranslation();
                vt.Export(emds, lcids, sheet, service, settings);
                StyleMutator.FontDefaults(sheet);
            }

            if ((settings.ExportForms || settings.ExportFormTabs || settings.ExportFormSections || settings.ExportFormFields) && emds.Count > 0)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting forms translations...");
                }

                var ft = new FormTranslation();

                ft.Export(emds, lcids, file.Workbook, service,
                    new FormExportOption
                    {
                        ExportForms = settings.ExportForms,
                        ExportFormTabs = settings.ExportFormTabs,
                        ExportFormSections = settings.ExportFormSections,
                        ExportFormFields = settings.ExportFormFields
                    }, settings);
            }

            if (settings.ExportSiteMap)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting SiteMap custom labels translations...");
                }

                var st = new SiteMapTranslation();

                st.Export(lcids, file.Workbook, service, settings);
            }

            if (settings.ExportDashboards)
            {
                if (worker != null && worker.WorkerReportsProgress)
                {
                    worker.ReportProgress(0, "Exporting Dashboards custom labels translations...");
                }

                var st = new DashboardTranslation();

                st.Export(lcids, file.Workbook, service, settings);
            }

            file.Save();
        }

        public void Import(string filePath, IOrganizationService service, BackgroundWorker worker = null)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var file = new ExcelPackage(stream);

                var emds = new List<EntityMetadata>();

                var forms = new List<Entity>();
                var dashboards = new List<Entity>();
                var ft = new FormTranslation();
                ft.Log += Engine_Log;
                ft.Result += Engine_OnResult;
                var st = new SiteMapTranslation();
                st.Log += Engine_Log;
                st.Result += Engine_OnResult;
                var db = new DashboardTranslation();
                db.Log += Engine_Log;
                db.Result += Engine_OnResult;
                bool hasFormContent = false;
                bool hasDashboardContent = false;
                bool hasSiteMapContent = false;

                var count = file.Workbook.Worksheets.Count(x =>
                    !x.Name.StartsWith("Forms ") && !x.Name.StartsWith("Dashboards ") &&
                    !x.Name.StartsWith("SiteMap "));

                var hasFormElts = file.Workbook.Worksheets.Any(x => x.Name.StartsWith("Forms "));
                var hasDashboardElts = file.Workbook.Worksheets.Any(x => x.Name.StartsWith("Dashboards "));
                var hasSiteMapElts = file.Workbook.Worksheets.Any(x => x.Name.StartsWith("SiteMap "));

                if (hasFormElts)
                {
                    count++;
                }

                if (hasDashboardElts)
                {
                    count++;
                }

                if (hasSiteMapElts)
                {
                    count++;
                }

                int overallProgress = 0;
                foreach (var sheet in file.Workbook.Worksheets)
                {
                    try
                    {
                        switch (sheet.Name)
                        {
                            case "Entities":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing entities translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var et = new EntityTranslation();
                                et.Result += Engine_OnResult;
                                et.Log += Engine_Log;
                                et.Import(sheet, emds, service, worker);

                                break;

                            case "Attributes":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing attributes translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var at = new AttributeTranslation();
                                at.Result += Engine_OnResult;
                                at.Log += Engine_Log;
                                at.Import(sheet, emds, service, worker);
                                break;

                            case "Relationships":
                                {
                                    worker.ReportProgressIfPossible(0, new ProgressInfo
                                    {
                                        Message = "Importing Relationships with custom label translations...",
                                        Item = 1,
                                        Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                    });

                                    var rt = new RelationshipTranslation();
                                    rt.Result += Engine_OnResult;
                                    rt.Log += Engine_Log;
                                    rt.Import(sheet, emds, service, worker);
                                    break;
                                }
                            case "RelationshipsNN":
                                {
                                    worker.ReportProgressIfPossible(0, new ProgressInfo
                                    {
                                        Message = "Importing NN Relationships with custom label translations...",
                                        Item = 1,
                                        Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                    });

                                    var rtNn = new RelationshipNnTranslation();
                                    rtNn.Result += Engine_OnResult;
                                    rtNn.Log += Engine_Log;
                                    rtNn.Import(sheet, emds, service, worker);
                                    break;
                                }
                            case "Global OptionSets":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing global optionsets translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var got = new GlobalOptionSetTranslation();
                                got.Result += Engine_OnResult;
                                got.Log += Engine_Log;
                                got.Import(sheet, service, worker);
                                break;

                            case "OptionSets":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing optionsets translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var ot = new OptionSetTranslation();
                                ot.Result += Engine_OnResult;
                                ot.Log += Engine_Log;
                                ot.Import(sheet, service, worker);
                                break;

                            case "Booleans":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing booleans translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var bt = new BooleanTranslation();
                                bt.Result += Engine_OnResult;
                                bt.Log += Engine_Log;
                                bt.Import(sheet, service, worker);
                                break;

                            case "Views":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing views translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var vt = new ViewTranslation();
                                vt.Result += Engine_OnResult;
                                vt.Log += Engine_Log;
                                vt.Import(sheet, service, worker);
                                break;

                            case "Charts":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing charts translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                var vt2 = new VisualizationTranslation();
                                vt2.Result += Engine_OnResult;
                                vt2.Log += Engine_Log;
                                vt2.Import(sheet, service, worker);
                                break;

                            case "Forms":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing forms translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                ft.ImportFormName(sheet, service, worker);
                                break;

                            case "Forms Tabs":
                                ft.PrepareFormTabs(sheet, service, forms);
                                hasFormContent = true;
                                break;

                            case "Forms Sections":
                                ft.PrepareFormSections(sheet, service, forms);
                                hasFormContent = true;
                                break;

                            case "Forms Fields":
                                ft.PrepareFormLabels(sheet, service, forms);
                                hasFormContent = true;
                                break;

                            case "Dashboards":
                                worker.ReportProgressIfPossible(0, new ProgressInfo
                                {
                                    Message = "Importing dashboard translations...",
                                    Item = 1,
                                    Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                                });

                                db.ImportFormName(sheet, service, worker);
                                break;

                            case "Dashboards Tabs":
                                db.PrepareFormTabs(sheet, service, dashboards);
                                hasDashboardContent = true;
                                break;

                            case "Dashboards Sections":
                                db.PrepareFormSections(sheet, service, dashboards);
                                hasDashboardContent = true;
                                break;

                            case "Dashboards Fields":
                                db.PrepareFormLabels(sheet, service, dashboards);
                                hasDashboardContent = true;
                                break;

                            case "SiteMap Areas":
                                st.PrepareAreas(sheet, service);
                                hasSiteMapContent = true;
                                break;

                            case "SiteMap Groups":
                                st.PrepareGroups(sheet, service);
                                hasSiteMapContent = true;
                                break;

                            case "SiteMap SubAreas":
                                st.PrepareSubAreas(sheet, service);
                                hasSiteMapContent = true;
                                break;
                        }
                    }
                    catch (Exception error)
                    {
                        // TODO handle this
                        //Engine_OnResult(this, new TranslationResultEventArgs
                        //{
                        //    Success = false,
                        //    SheetName = sheet.Name,
                        //    Message = error.Message
                        //});
                    }
                    finally
                    {
                        overallProgress++;
                    }
                }

                if (hasFormContent)
                {
                    worker.ReportProgressIfPossible(0, new ProgressInfo
                    {
                        Message = "Importing form content translations...",
                        Item = 1,
                        Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                    });

                    ft.ImportFormsContent(service, forms, worker);
                }

                if (hasDashboardContent)
                {
                    worker.ReportProgressIfPossible(0, new ProgressInfo
                    {
                        Message = "Importing dashboard content translations...",
                        Item = 1,
                        Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                    });

                    db.ImportFormsContent(service, dashboards, worker);
                }

                if (hasSiteMapContent)
                {
                    worker.ReportProgressIfPossible(0, new ProgressInfo
                    {
                        Message = "Importing SiteMap translations...",
                        Item = 1,
                        Overall = overallProgress == 0 ? 1 : overallProgress * 100 / count
                    });

                    st.Import(service);
                }

                OnProgress?.Invoke(this, new TranslationProgressEventArgs());

                worker.ReportProgressIfPossible(0, new ProgressInfo
                {
                    Message = "Publishing customizations...",
                    Item = 100,
                    Overall = 100
                });

                var paxRequest = new PublishAllXmlRequest();
                service.Execute(paxRequest);

                worker.ReportProgressIfPossible(0, new ProgressInfo
                {
                    Message = "",
                    Item = 100,
                    Overall = 100
                });
            }
        }

        private void Engine_Log(object sender, LogEventArgs e)
        {
            OnLog?.Invoke(sender, e);
        }

        private void Engine_OnResult(object sender, TranslationProgressEventArgs e)
        {
            OnProgress?.Invoke(sender, e);
        }
    }
}