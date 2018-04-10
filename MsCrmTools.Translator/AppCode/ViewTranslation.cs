using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace MsCrmTools.Translator.AppCode
{
    public class ViewTranslation : BaseTranslation
    {
        /// <summary>
        ///
        /// </summary>
        /// <example>
        /// viewId;entityLogicalName;viewName;ViewType;Type;LCID1;LCID2;...;LCODX
        /// </example>
        /// <param name="entities"></param>
        /// <param name="languages"></param>
        /// <param name="sheet"></param>
        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorksheet sheet, IOrganizationService service, ExportSettings settings)
        {
            var line = 0;
            var cell = 0;

            AddHeader(sheet, languages);

            var crmViews = new List<CrmView>();

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                if (!entity.MetadataId.HasValue)
                    continue;

                var views = RetrieveViews(entity.LogicalName, entity.ObjectTypeCode.Value, service);

                foreach (var view in views)
                {
                    var crmView = crmViews.FirstOrDefault(cv => cv.Id == view.Id);
                    if (crmView == null)
                    {
                        crmView = new CrmView
                        {
                            Id = view.Id,
                            Entity = view.GetAttributeValue<string>("returnedtypecode"),
                            Type = view.GetAttributeValue<int>("querytype"),
                            Names = new Dictionary<int, string>(),
                            Descriptions = new Dictionary<int, string>()
                        };
                        crmViews.Add(crmView);
                    }

                    RetrieveLocLabelsRequest request;
                    RetrieveLocLabelsResponse response;

                    if (settings.ExportNames)
                    {
                        // Names
                        request = new RetrieveLocLabelsRequest
                        {
                            AttributeName = "name",
                            EntityMoniker = new EntityReference("savedquery", view.Id)
                        };

                        response = (RetrieveLocLabelsResponse)service.Execute(request);
                        foreach (var locLabel in response.Label.LocalizedLabels)
                        {
                            crmView.Names.Add(locLabel.LanguageCode, locLabel.Label);
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        // Descriptions
                        request = new RetrieveLocLabelsRequest
                        {
                            AttributeName = "description",
                            EntityMoniker = new EntityReference("savedquery", view.Id)
                        };

                        response = (RetrieveLocLabelsResponse)service.Execute(request);
                        foreach (var locLabel in response.Label.LocalizedLabels)
                        {
                            crmView.Descriptions.Add(locLabel.LanguageCode, locLabel.Label);
                        }
                    }
                }
            }

            foreach (var crmView in crmViews.OrderBy(cv => cv.Entity).ThenBy(cv => cv.Type))
            {
                if (settings.ExportNames)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmView.Id.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmView.Entity;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmView.Type;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Name";

                    foreach (var lcid in languages)
                    {
                        var name = crmView.Names.FirstOrDefault(n => n.Key == lcid);
                        if (name.Value != null)
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = name.Value;
                        else
                        {
                            cell++;
                        }
                    }
                }

                if (settings.ExportDescriptions)
                {
                    line++;
                    cell = 0;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmView.Id.ToString("B");
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmView.Entity;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = crmView.Type;
                    ZeroBasedSheet.Cell(sheet, line, cell++).Value = "Description";

                    foreach (var lcid in languages)
                    {
                        var desc = crmView.Descriptions.FirstOrDefault(n => n.Key == lcid);
                        if (desc.Value != null)
                            ZeroBasedSheet.Cell(sheet, line, cell++).Value = desc.Value;
                        else
                        {
                            cell++;
                        }
                    }
                }
            }

            // Applying style to cells
            for (int i = 0; i < (4 + languages.Count); i++)
            {
                StyleMutator.TitleCell(ZeroBasedSheet.Cell(sheet, 0, i).Style);
            }

            for (int i = 1; i <= line; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(sheet, i, j).Style);
                }
            }
        }

        public void Import(ExcelWorksheet sheet, IOrganizationService service, BackgroundWorker worker)
        {
            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            var requests = new List<SetLocLabelsRequest>();
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                var currentViewId = new Guid(ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString());
                var request = new SetLocLabelsRequest
                {
                    EntityMoniker = new EntityReference("savedquery", currentViewId),
                    AttributeName = ZeroBasedSheet.Cell(sheet, rowI, 3).Value.ToString() == "Name" ? "name" : "description"
                };

                var labels = new List<LocalizedLabel>();

                var columnIndex = 4;
                while (columnIndex < cellsCount)
                {
                    if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                        labels.Add(new LocalizedLabel(label, lcid));
                    }

                    columnIndex++;
                }

                request.Labels = labels.ToArray();

                requests.Add(request);
            }

            int i = 0;
            foreach (var request in requests)
            {
                try
                {
                    service.Execute(request);

                    OnResult(new TranslationResultEventArgs
                    {
                        Success = true,
                        SheetName = sheet.Name
                    });
                }
                catch (Exception error)
                {
                    OnResult(new TranslationResultEventArgs
                    {
                        Success = false,
                        SheetName = sheet.Name,
                        Message = $"{request.EntityMoniker.Id}/{request.AttributeName}: {error.Message}"
                    });
                }

                i++;
                worker.ReportProgressIfPossible(0, new ProgressInfo
                {
                    Item = i * 100 / requests.Count
                });
            }
        }

        private void AddHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "View Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "ViewType";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private List<Entity> RetrieveViews(string entityLogicalName, int objectTypeCode, IOrganizationService service)
        {
            try
            {
                var qba = new QueryByAttribute
                {
                    EntityName = "savedquery",
                };

                qba.Attributes.Add("returnedtypecode");
                qba.Values.Add(objectTypeCode);

                EntityCollection views = service.RetrieveMultiple(qba);

                return views.Entities.ToList();
            }
            catch (Exception error)
            {
                throw new Exception("Error while retrieving views: " + error.Message);
            }
        }
    }
}