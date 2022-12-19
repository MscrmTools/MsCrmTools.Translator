using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml;
using ExcelWorksheet = OfficeOpenXml.ExcelWorksheet;

namespace MsCrmTools.Translator.AppCode
{
    public class FormTranslation : BaseTranslation
    {
        private static ExportSettings settings;

        public FormTranslation()
        {
            name = "Forms";
        }

        public void Export(List<EntityMetadata> entities, List<int> languages, ExcelWorkbook file, IOrganizationService service, FormExportOption options, ExportSettings esettings)
        {
            settings = esettings;

            // Retrieve current user language information
            var setting = GetCurrentUserSettings(service);

            var userSettingLcid = setting.GetAttributeValue<int>("uilanguageid");
            if (userSettingLcid == 0)
            {
                userSettingLcid = languages.First();
            }

            var currentSetting = userSettingLcid;

            var crmForms = new List<CrmForm>();
            var crmFormTabs = new List<CrmFormTab>();
            var crmFormSections = new List<CrmFormSection>();
            var crmFormLabels = new List<CrmFormLabel>();

            foreach (var lcid in languages)
            {
                if (currentSetting != lcid)
                {
                    setting["localeid"] = lcid;
                    setting["uilanguageid"] = lcid;
                    setting["helplanguageid"] = lcid;
                    service.Update(setting);
                    currentSetting = lcid;

                    Thread.Sleep(2000);
                }

                foreach (var entity in entities.OrderBy(e => e.LogicalName))
                {
                    if (!entity.MetadataId.HasValue)
                        continue;

                    var forms = RetrieveEntityFormList(entity.LogicalName, service);

                    foreach (var form in forms)
                    {
                        #region Tabs

                        if (options.ExportFormTabs || options.ExportFormSections || options.ExportFormFields)
                        {
                            // Load Xml definition of form
                            var sFormXml = form.GetAttributeValue<string>("formxml");
                            var formXml = new XmlDocument();
                            formXml.LoadXml(sFormXml);

                            // Specific for header
                            if (options.ExportFormFields)
                            {
                                var cellNodes = formXml.DocumentElement.SelectNodes("header/rows/row/cell");
                                foreach (XmlNode cellNode in cellNodes)
                                {
                                    ExtractField(cellNode, crmFormLabels, form, null, null, entity, lcid);
                                }
                            }

                            foreach (XmlNode tabNode in formXml.SelectNodes("//tab"))
                            {
                                var tabName = ExtractTabName(tabNode, lcid, crmFormTabs, form, entity);

                                #region Sections

                                if (options.ExportFormSections || options.ExportFormFields)
                                {
                                    foreach (
                                        XmlNode sectionNode in tabNode.SelectNodes("columns/column/sections/section"))
                                    {
                                        var sectionName = ExtractSection(sectionNode, lcid, crmFormSections, form,
                                            tabName, entity);

                                        #region Labels

                                        if (options.ExportFormFields)
                                        {
                                            foreach (XmlNode labelNode in sectionNode.SelectNodes("rows/row/cell"))
                                            {
                                                ExtractField(labelNode, crmFormLabels, form, tabName, sectionName,
                                                    entity, lcid);
                                            }
                                        }

                                        #endregion Labels
                                    }
                                }

                                #endregion Sections
                            }

                            // Specific for footer
                            if (options.ExportFormFields)
                            {
                                var cellNodes = formXml.DocumentElement.SelectNodes("footer/rows/row/cell");
                                foreach (XmlNode cellNode in cellNodes)
                                {
                                    ExtractField(cellNode, crmFormLabels, form, null, null, entity, lcid);
                                }
                            }
                        }

                        #endregion Tabs
                    }
                }
            }

            if (userSettingLcid != currentSetting)
            {
                setting["localeid"] = userSettingLcid;
                setting["uilanguageid"] = userSettingLcid;
                setting["helplanguageid"] = userSettingLcid;
                service.Update(setting);
            }

            foreach (var entity in entities.OrderBy(e => e.LogicalName))
            {
                if (!entity.MetadataId.HasValue)
                    continue;

                var forms = RetrieveEntityFormList(entity.LogicalName, service);

                foreach (var form in forms)
                {
                    var crmForm =
                        crmForms.FirstOrDefault(f => f.FormUniqueId == form.GetAttributeValue<Guid>("formidunique"));
                    if (crmForm == null)
                    {
                        crmForm = new CrmForm
                        {
                            FormUniqueId = form.GetAttributeValue<Guid>("formidunique"),
                            Id = form.GetAttributeValue<Guid>("formid"),
                            Entity = entity.LogicalName,
                            Names = new Dictionary<int, string>(),
                            Descriptions = new Dictionary<int, string>(),
                            Type = form.FormattedValues["type"]
                        };
                        crmForms.Add(crmForm);
                    }

                    RetrieveLocLabelsRequest request;
                    RetrieveLocLabelsResponse response;

                    if (settings.ExportNames)
                    {
                        // Names
                        request = new RetrieveLocLabelsRequest
                        {
                            AttributeName = "name",
                            EntityMoniker = new EntityReference("systemform", form.Id)
                        };

                        response = (RetrieveLocLabelsResponse)service.Execute(request);
                        foreach (var locLabel in response.Label.LocalizedLabels)
                        {
                            crmForm.Names.Add(locLabel.LanguageCode, locLabel.Label);
                        }
                    }

                    if (settings.ExportDescriptions)
                    {
                        // Descriptions
                        request = new RetrieveLocLabelsRequest
                        {
                            AttributeName = "description",
                            EntityMoniker = new EntityReference("systemform", form.Id)
                        };

                        response = (RetrieveLocLabelsResponse)service.Execute(request);
                        foreach (var locLabel in response.Label.LocalizedLabels)
                        {
                            crmForm.Descriptions.Add(locLabel.LanguageCode, locLabel.Label);
                        }
                    }
                }
            }

            var line = 0;
            if (options.ExportForms)
            {
                var formSheet = file.Worksheets.Add("Forms");
                AddFormHeader(formSheet, languages);

                foreach (var crmForm in crmForms)
                {
                    line = ExportForm(languages, formSheet, line, crmForm);
                }

                // Applying style to cells
                for (int i = 0; i < (4 + languages.Count); i++)
                {
                    StyleMutator.TitleCell(ZeroBasedSheet.Cell(formSheet, 0, i).Style);
                }

                for (int i = 1; i <= line; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(formSheet, i, j).Style);
                    }
                }
            }

            if (options.ExportFormTabs)
            {
                var tabSheet = file.Worksheets.Add("Forms Tabs");
                line = 1;
                AddFormTabHeader(tabSheet, languages);
                foreach (var crmFormTab in crmFormTabs)
                {
                    line = ExportTab(languages, tabSheet, line, crmFormTab);
                }

                // Applying style to cells
                for (int i = 0; i < (5 + languages.Count); i++)
                {
                    StyleMutator.TitleCell(ZeroBasedSheet.Cell(tabSheet, 0, i).Style);
                }

                for (int i = 1; i < line; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(tabSheet, i, j).Style);
                    }
                }
            }

            if (options.ExportFormSections)
            {
                var sectionSheet = file.Worksheets.Add("Forms Sections");
                line = 1;
                AddFormSectionHeader(sectionSheet, languages);
                foreach (var crmFormSection in crmFormSections)
                {
                    line = ExportSection(languages, sectionSheet, line, crmFormSection);
                }

                // Applying style to cells
                for (int i = 0; i < (6 + languages.Count); i++)
                {
                    StyleMutator.TitleCell(ZeroBasedSheet.Cell(sectionSheet, 0, i).Style);
                }

                for (int i = 1; i < line; i++)
                {
                    for (int j = 0; j < 6; j++)
                    {
                        StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(sectionSheet, i, j).Style);
                    }
                }
            }

            if (options.ExportFormFields)
            {
                var labelSheet = file.Worksheets.Add("Forms Fields");
                AddFormLabelsHeader(labelSheet, languages);
                line = 1;
                foreach (var crmFormLabel in crmFormLabels)
                {
                    line = ExportField(languages, labelSheet, line, crmFormLabel);
                }

                // Applying style to cells
                for (int i = 0; i < (8 + languages.Count); i++)
                {
                    StyleMutator.TitleCell(ZeroBasedSheet.Cell(labelSheet, 0, i).Style);
                }

                for (int i = 1; i < line; i++)
                {
                    for (int j = 0; j < 8; j++)
                    {
                        StyleMutator.HighlightedCell(ZeroBasedSheet.Cell(labelSheet, i, j).Style);
                    }
                }
            }
        }

        public void ImportFormName(ExcelWorksheet sheet, IOrganizationService service, BackgroundWorker worker)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            var requests = new List<SetLocLabelsRequest>();
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                var currentFormId = new Guid(ZeroBasedSheet.Cell(sheet, rowI, 1).Value.ToString());

                var locLabel = ((RetrieveLocLabelsResponse)service.Execute(new RetrieveLocLabelsRequest
                {
                    EntityMoniker = new EntityReference("systemform", currentFormId),
                    AttributeName = ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString() == "Name" ? "name" : "description"
                })).Label;

                var labels = locLabel.LocalizedLabels.ToList();

                var request = new SetLocLabelsRequest
                {
                    EntityMoniker = new EntityReference("systemform", currentFormId),
                    AttributeName = ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString() == "Name" ? "name" : "description",
                    Labels = locLabel.LocalizedLabels.ToArray()
                };

                var columnIndex = 5;
                while (columnIndex < cellsCount)
                {
                    if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                    {
                        var lcid = int.Parse(ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString());
                        var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                        var translatedLabel = labels.FirstOrDefault(x => x.LanguageCode == lcid);
                        if (translatedLabel == null)
                        {
                            translatedLabel = new LocalizedLabel(label, lcid);
                            labels.Add(translatedLabel);
                        }
                        else
                        {
                            translatedLabel.Label = label;
                        }
                    }

                    columnIndex++;
                }

                request.Labels = labels.ToArray();

                requests.Add(request);
            }

            OnLog(new LogEventArgs($"Importing {sheet.Name} translations"));

            var setting = GetCurrentUserSettings(service);
            var userSettingLcid = setting.GetAttributeValue<int>("uilanguageid");
            var currentSetting = userSettingLcid;

            int orgLcid = GetCurrentOrgBaseLanguage(service);
            if (currentSetting != orgLcid)
            {
                setting["localeid"] = orgLcid;
                setting["uilanguageid"] = orgLcid;
                setting["helplanguageid"] = orgLcid;
                service.Update(setting);
                currentSetting = orgLcid;

                Thread.Sleep(2000);
            }

            var arg = new TranslationProgressEventArgs { SheetName = sheet.Name };
            foreach (var request in requests)
            {
                AddRequest(request);
                ExecuteMultiple(service, arg, requests.Count);
            }
            ExecuteMultiple(service, arg, requests.Count, true);

            if (currentSetting != userSettingLcid)
            {
                setting["localeid"] = userSettingLcid;
                setting["uilanguageid"] = userSettingLcid;
                setting["helplanguageid"] = userSettingLcid;
                service.Update(setting);

                Thread.Sleep(2000);
            }
        }

        public void ImportFormsContent(IOrganizationService service, List<Entity> forms, BackgroundWorker worker)
        {
            name = "Forms Contents";

            OnLog(new LogEventArgs($"Importing {name} translations"));

            var setting = GetCurrentUserSettings(service);
            var userSettingLcid = setting.GetAttributeValue<int>("uilanguageid");
            var currentSetting = userSettingLcid;

            var response = (RetrieveAvailableLanguagesResponse)service.Execute(new RetrieveAvailableLanguagesRequest());
            foreach (var lcid in response.LocaleIds)
            {
                if (currentSetting != lcid)
                {
                    setting["localeid"] = lcid;
                    setting["uilanguageid"] = lcid;
                    setting["helplanguageid"] = lcid;
                    service.Update(setting);
                    currentSetting = lcid;

                    OnLog(new LogEventArgs($"Current user language changed to {currentSetting}"));
                }

                OnLog(new LogEventArgs($"Importing translations for language {currentSetting}"));

                var arg = new TranslationProgressEventArgs();
                foreach (var form in forms)
                {
                    AddRequest(new UpdateRequest { Target = form });
                    ExecuteMultiple(service, arg, forms.Count);
                }
                ExecuteMultiple(service, arg, forms.Count, true);
            }

            if (currentSetting != userSettingLcid)
            {
                setting["localeid"] = userSettingLcid;
                setting["uilanguageid"] = userSettingLcid;
                setting["helplanguageid"] = userSettingLcid;
                service.Update(setting);
                currentSetting = userSettingLcid;
                OnLog(new LogEventArgs($"Current user language changed to {currentSetting}"));
            }
        }

        public void PrepareFormLabels(ExcelWorksheet sheet, IOrganizationService service, List<Entity> forms)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 7)) continue;

                var labelId = ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString();
                var formId = new Guid(ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString());

                var form = forms.FirstOrDefault(f => f.Id == formId);
                if (form == null)
                {
                    try
                    {
                        form = service.Retrieve("systemform", formId, new ColumnSet(new[] { "formxml" }));
                        forms.Add(form);
                    }
                    catch (Exception error) //lets not fail if the form is no more available in CRM
                    {
                        OnLog(new LogEventArgs($"{sheet.Name}: {formId}: {error.Message}"));

                        continue;   //form is not found so no need to process further.
                    }
                }

                // Load formxml
                var formXml = form.GetAttributeValue<string>("formxml");
                var docXml = new XmlDocument();
                docXml.LoadXml(formXml);

                var cellNode =
                    docXml.DocumentElement.SelectSingleNode(
                        //tabs/tab/columns/column/sections/section/rows/row
                        string.Format("//cell[translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ{{}}','abcdefghijklmnopqrstuvwxyz')='{0}']", new Guid(labelId).ToString()));
                if (cellNode != null)
                {
                    var columnIndex = 8;
                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString();
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            UpdateXmlNode(cellNode, lcid, label);
                        }

                        columnIndex++;
                    }
                }

                form["formxml"] = docXml.OuterXml;
            }
        }

        public void PrepareFormSections(ExcelWorksheet sheet, IOrganizationService service, List<Entity> forms)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 5)) continue;

                var sectionId = ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString();
                var formId = new Guid(ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString());

                var form = forms.FirstOrDefault(f => f.Id == formId);
                if (form == null)
                {
                    try
                    {
                        form = service.Retrieve("systemform", formId, new ColumnSet("formxml"));
                        forms.Add(form);
                    }
                    catch (Exception error) //lets not fail if the form is no more available in CRM
                    {
                        OnLog(new LogEventArgs
                        {
                            Type = LogType.Warning,
                            Message = $"Cannot find form {formId}: {error.Message}"
                        });

                        continue;   //form is not found so no need to process further.
                    }
                }

                // Load formxml
                var formXml = form.GetAttributeValue<string>("formxml");
                var docXml = new XmlDocument();
                docXml.LoadXml(formXml);

                var sectionNode =
                    docXml.DocumentElement.SelectSingleNode(
                        string.Format("tabs/tab/columns/column/sections/section[translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ{{}}','abcdefghijklmnopqrstuvwxyz')='{0}']", new Guid(sectionId).ToString()));
                if (sectionNode != null)
                {
                    var columnIndex = 6;
                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString();
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            UpdateXmlNode(sectionNode, lcid, label);
                        }

                        columnIndex++;
                    }
                }

                form["formxml"] = docXml.OuterXml;
            }
        }

        public void PrepareFormTabs(ExcelWorksheet sheet, IOrganizationService service, List<Entity> forms)
        {
            OnLog(new LogEventArgs($"Reading {sheet.Name}"));

            var rowsCount = sheet.Dimension.Rows;
            var cellsCount = sheet.Dimension.Columns;
            for (var rowI = 1; rowI < rowsCount; rowI++)
            {
                if (HasEmptyCells(sheet, rowI, 4)) continue;

                var tabId = ZeroBasedSheet.Cell(sheet, rowI, 0).Value.ToString();
                var formId = new Guid(ZeroBasedSheet.Cell(sheet, rowI, 4).Value.ToString());

                var form = forms.FirstOrDefault(f => f.Id == formId);
                if (form == null)
                {
                    try
                    {
                        form = service.Retrieve("systemform", formId, new ColumnSet(new[] { "formxml" }));
                        forms.Add(form);
                    }
                    catch (Exception error) //lets not fail if the form is no more available in CRM
                    {
                        OnLog(new LogEventArgs($"{sheet.Name}: {formId}: {error.Message}"));

                        continue;   //form is not found so no need to process further.
                    }
                }

                // Load formxml
                var formXml = form.GetAttributeValue<string>("formxml");
                var docXml = new XmlDocument();
                docXml.LoadXml(formXml);

                var tabNode = docXml.DocumentElement.SelectSingleNode(string.Format("tabs/tab[translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ{{}}','abcdefghijklmnopqrstuvwxyz')='{0}']", new Guid(tabId).ToString()));
                if (tabNode != null)
                {
                    var columnIndex = 5;
                    while (columnIndex < cellsCount)
                    {
                        if (ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value != null)
                        {
                            var lcid = ZeroBasedSheet.Cell(sheet, 0, columnIndex).Value.ToString();
                            var label = ZeroBasedSheet.Cell(sheet, rowI, columnIndex).Value.ToString();

                            UpdateXmlNode(tabNode, lcid, label);
                        }

                        columnIndex++;
                    }
                }

                form["formxml"] = docXml.OuterXml;
            }
        }

        private static int ExportField(List<int> languages, ExcelWorksheet labelSheet, int line,
            CrmFormLabel crmFormLabel)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.Id.ToString("B");
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.Entity;
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.Form;
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.FormUniqueId.ToString("B");
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.FormId.ToString("B");
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.Tab;
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.Section;
            ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = crmFormLabel.Attribute;

            foreach (var lcid in languages)
            {
                bool exists = crmFormLabel.Names.ContainsKey(lcid);
                ZeroBasedSheet.Cell(labelSheet, line, cell++).Value = exists
                    ? crmFormLabel.Names.First(n => n.Key == lcid).Value
                    : string.Empty;
            }

            line++;
            return line;
        }

        private static int ExportForm(List<int> languages, ExcelWorksheet formSheet, int line, CrmForm crmForm)
        {
            int cell;

            if (settings.ExportNames)
            {
                line++;
                cell = 0;

                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.FormUniqueId.ToString("B");
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.Id.ToString("B");
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.Entity;
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.Type;
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = "Name";

                foreach (var lcid in languages)
                {
                    var name = crmForm.Names.FirstOrDefault(n => n.Key == lcid);
                    if (name.Value != null)
                        ZeroBasedSheet.Cell(formSheet, line, cell++).Value = name.Value;
                    else
                        cell++;
                }
            }

            if (settings.ExportDescriptions)
            {
                line++;
                cell = 0;

                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.FormUniqueId.ToString("B");
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.Id.ToString("B");
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.Entity;
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = crmForm.Type;
                ZeroBasedSheet.Cell(formSheet, line, cell++).Value = "Description";

                foreach (var lcid in languages)
                {
                    var desc = crmForm.Descriptions.FirstOrDefault(n => n.Key == lcid);
                    if (desc.Value != null)
                        ZeroBasedSheet.Cell(formSheet, line, cell++).Value = desc.Value;
                    else
                        cell++;
                }
            }

            return line;
        }

        private static int ExportSection(List<int> languages, ExcelWorksheet sectionSheet, int line,
            CrmFormSection crmFormSection)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = crmFormSection.Id.ToString("B");
            ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = crmFormSection.Entity;
            ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = crmFormSection.Form;
            ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = crmFormSection.FormUniqueId.ToString("B");
            ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = crmFormSection.FormId.ToString("B");
            ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = crmFormSection.Tab;

            foreach (var lcid in languages)
            {
                bool exists = crmFormSection.Names.ContainsKey(lcid);
                ZeroBasedSheet.Cell(sectionSheet, line, cell++).Value = exists
                    ? crmFormSection.Names.First(n => n.Key == lcid).Value
                    : string.Empty;
            }

            line++;
            return line;
        }

        private void AddFormHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Unique Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Type";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Type";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void AddFormLabelsHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Label Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Unique Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Tab Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Section Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Attribute";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void AddFormSectionHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Section Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Unique Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Tab Name";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void AddFormTabHeader(ExcelWorksheet sheet, IEnumerable<int> languages)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Tab Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Entity Logical Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Name";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Unique Id";
            ZeroBasedSheet.Cell(sheet, 0, cell++).Value = "Form Id";

            foreach (var lcid in languages)
            {
                ZeroBasedSheet.Cell(sheet, 0, cell++).Value = lcid.ToString(CultureInfo.InvariantCulture);
            }
        }

        private int ExportTab(List<int> languages, ExcelWorksheet tabSheet, int line, CrmFormTab crmFormTab)
        {
            var cell = 0;

            ZeroBasedSheet.Cell(tabSheet, line, cell++).Value = crmFormTab.Id.ToString("B");
            ZeroBasedSheet.Cell(tabSheet, line, cell++).Value = crmFormTab.Entity;
            ZeroBasedSheet.Cell(tabSheet, line, cell++).Value = crmFormTab.Form;
            ZeroBasedSheet.Cell(tabSheet, line, cell++).Value = crmFormTab.FormUniqueId.ToString("B");
            ZeroBasedSheet.Cell(tabSheet, line, cell++).Value = crmFormTab.FormId.ToString("B");

            foreach (var lcid in languages)
            {
                bool exists = crmFormTab.Names.ContainsKey(lcid);
                ZeroBasedSheet.Cell(tabSheet, line, cell++).Value = exists
                    ? crmFormTab.Names.First(n => n.Key == lcid).Value
                    : string.Empty;
            }

            line++;
            return line;
        }

        private void ExtractField(XmlNode cellNode, List<CrmFormLabel> crmFormLabels, Entity form, string tabName,
            string sectionName, EntityMetadata entity, int lcid)
        {
            if (cellNode.Attributes == null)
                return;

            var cellIdAttr = cellNode.Attributes["id"];
            if (cellIdAttr == null)
                return;

            if (cellNode.ChildNodes.Count == 0)
                return;

            var controlNode = cellNode.SelectSingleNode("control");
            if (controlNode == null || controlNode.Attributes == null)
                return;

            //var crmFormField = crmFormLabels.FirstOrDefault(f => f.Id == new Guid(cellIdAttr.Value) && f.FormId == form.Id);
            var crmFormField =
                crmFormLabels.FirstOrDefault(
                    f =>
                        f.Id == new Guid(cellIdAttr.Value) &&
                        f.FormUniqueId == form.GetAttributeValue<Guid>("formidunique"));
            if (crmFormField == null)
            {
                crmFormField = new CrmFormLabel
                {
                    Id = new Guid(cellIdAttr.Value),
                    Form = form.GetAttributeValue<string>("name"),
                    FormUniqueId = form.GetAttributeValue<Guid>("formidunique"),
                    FormId = form.GetAttributeValue<Guid>("formid"),
                    Tab = tabName,
                    Section = sectionName,
                    Entity = entity.LogicalName,
                    Attribute = controlNode.Attributes["id"].Value,
                    Names = new Dictionary<int, string>()
                };
                crmFormLabels.Add(crmFormField);
            }

            var labelNode = cellNode.SelectSingleNode("labels/label[@languagecode='" + lcid + "']");
            var labelNodeAttributes = labelNode?.Attributes;
            var labelDescription = labelNodeAttributes?["description"];

            if (crmFormField.Names.ContainsKey(lcid))
            {
                return;
            }

            crmFormField.Names.Add(lcid, labelDescription == null ? string.Empty : labelDescription.Value);
        }

        private string ExtractSection(XmlNode sectionNode, int lcid, List<CrmFormSection> crmFormSections, Entity form,
            string tabName, EntityMetadata entity)
        {
            if (sectionNode.Attributes == null || sectionNode.Attributes["id"] == null)
                return string.Empty;
            var sectionId = sectionNode.Attributes["id"].Value;

            var sectionLabelNode = sectionNode.SelectSingleNode("labels/label[@languagecode='" + lcid + "']");
            if (sectionLabelNode == null || sectionLabelNode.Attributes == null)
                return string.Empty;

            var sectionNameAttr = sectionLabelNode.Attributes["description"];
            if (sectionNameAttr == null)
                return string.Empty;
            var sectionName = sectionNameAttr.Value;

            var crmFormSection =
                crmFormSections.FirstOrDefault(
                    f => f.Id == new Guid(sectionId) && f.FormUniqueId == form.GetAttributeValue<Guid>("formidunique"));
            if (crmFormSection == null)
            {
                crmFormSection = new CrmFormSection
                {
                    Id = new Guid(sectionId),
                    FormUniqueId = form.GetAttributeValue<Guid>("formidunique"),
                    FormId = form.GetAttributeValue<Guid>("formid"),
                    Form = form.GetAttributeValue<string>("name"),
                    Tab = tabName,
                    Entity = entity.LogicalName,
                    Names = new Dictionary<int, string>()
                };
                crmFormSections.Add(crmFormSection);
            }
            if (crmFormSection.Names.ContainsKey(lcid))
            {
                return sectionName;
            }
            crmFormSection.Names.Add(lcid, sectionName);
            return sectionName;
        }

        private string ExtractTabName(XmlNode tabNode, int lcid, List<CrmFormTab> crmFormTabs, Entity form,
            EntityMetadata entity)
        {
            if (tabNode.Attributes == null || tabNode.Attributes["id"] == null)
                return string.Empty;

            var tabId = tabNode.Attributes["id"].Value;

            var tabLabelNode = tabNode.SelectSingleNode("labels/label[@languagecode='" + lcid + "']");
            if (tabLabelNode == null || tabLabelNode.Attributes == null)
                return string.Empty;

            var tabLabelDescAttr = tabLabelNode.Attributes["description"];
            if (tabLabelDescAttr == null)
                return string.Empty;

            var tabName = tabLabelDescAttr.Value;

            var crmFormTab =
                crmFormTabs.FirstOrDefault(
                    f => f.Id == new Guid(tabId) && f.FormUniqueId == form.GetAttributeValue<Guid>("formidunique"));
            if (crmFormTab == null)
            {
                crmFormTab = new CrmFormTab
                {
                    Id = new Guid(tabId),
                    FormUniqueId = form.GetAttributeValue<Guid>("formidunique"),
                    FormId = form.GetAttributeValue<Guid>("formid"),
                    Form = form.GetAttributeValue<string>("name"),
                    Entity = entity.LogicalName,
                    Names = new Dictionary<int, string>()
                };
                crmFormTabs.Add(crmFormTab);
            }

            if (crmFormTab.Names.ContainsKey(lcid))
            {
                return tabName;
            }

            crmFormTab.Names.Add(lcid, tabName);
            return tabName;
        }

        private int GetCurrentOrgBaseLanguage(IOrganizationService service)
        {
            var qe = new QueryExpression("organization");
            qe.ColumnSet = new ColumnSet(new[] { "languagecode" });
            var settings = service.RetrieveMultiple(qe);

            return settings[0].GetAttributeValue<int>("languagecode");
        }

        private Entity GetCurrentUserSettings(IOrganizationService service)
        {
            var qe = new QueryExpression("usersettings");
            qe.ColumnSet = new ColumnSet(new[] { "uilanguageid", "localeid" });
            qe.Criteria = new FilterExpression();
            qe.Criteria.AddCondition("systemuserid", ConditionOperator.EqualUserId);
            var settings = service.RetrieveMultiple(qe);

            return settings[0];
        }

        private IEnumerable<Entity> RetrieveEntityFormList(string logicalName, IOrganizationService oService)
        {
            var qe = new QueryExpression("systemform")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("objecttypecode", ConditionOperator.Equal, logicalName),
                        new ConditionExpression("type", ConditionOperator.In, new[] {2, 6, 7})
                    }
                }
            };

            var ec = oService.RetrieveMultiple(qe);

            return ec.Entities;
        }

        private void UpdateXmlNode(XmlNode node, string lcid, string description)
        {
            var labelsNode = node.SelectSingleNode("labels");
            if (labelsNode == null)
            {
                labelsNode = node.OwnerDocument.CreateElement("labels");
                node.AppendChild(labelsNode);
            }

            var labelNode = labelsNode.SelectSingleNode(string.Format("label[@languagecode='{0}']", lcid));
            if (labelNode == null)
            {
                labelNode = node.OwnerDocument.CreateElement("label");
                labelsNode.AppendChild(labelNode);

                var languageAttr = node.OwnerDocument.CreateAttribute("languagecode");
                languageAttr.Value = lcid;
                labelNode.Attributes.Append(languageAttr);
                var descriptionAttr = node.OwnerDocument.CreateAttribute("description");
                labelNode.Attributes.Append(descriptionAttr);
            }

            labelNode.Attributes["description"].Value = description;
        }
    }
}