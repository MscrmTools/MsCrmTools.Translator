using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Linq;

namespace MsCrmTools.Translator.AppCode
{
    public enum LogType
    {
        Info,
        Warning,
        Error
    }

    public class BaseTranslation
    {
        public static int BulkCount;
        protected string name;
        private ExecuteMultipleRequest request;

        public event EventHandler<LogEventArgs> Log;

        public event EventHandler<TranslationProgressEventArgs> Result;

        public virtual void OnLog(LogEventArgs e)
        {
            Log?.Invoke(this, e);
        }

        public virtual void OnResult(TranslationProgressEventArgs e)
        {
            Result?.Invoke(this, e);
        }

        protected void AddRequest(OrganizationRequest or)
        {
            if (request == null) InitMultipleRequest();

            request.Requests.Add(or);
        }

        protected void ExecuteMultiple(IOrganizationService service, TranslationProgressEventArgs e, int total, bool forceUpdate = false)
        {
            if (request == null) return;
            if (request.Requests.Count % BulkCount != 0 && !forceUpdate) return;

            e.SheetName = name;
            e.TotalItems = total;

            OnResult(e);

            var bulkResponse = (ExecuteMultipleResponse)service.Execute(request);

            if (bulkResponse.IsFaulted)
            {
                e.FailureCount += bulkResponse.Responses.Count(r => r.Fault != null);
                e.SuccessCount += request.Requests.Count - bulkResponse.Responses.Count;

                foreach (var response in bulkResponse.Responses)
                {
                    if (response.Fault != null)
                    {
                        string message;
                        var faultIndex = response.RequestIndex;
                        var faultRequest = request.Requests[faultIndex];

                        if (faultRequest is UpdateRequest ur)
                        {
                            message =
                                $"Error while updating record {ur.Target.LogicalName} ({ur.Target.Id}): {response.Fault.Message}";
                        }
                        else if (faultRequest is UpdateAttributeRequest uar)
                        {
                            message =
                                $"Error while updating attribute {uar.Attribute.LogicalName}: {response.Fault.Message}";
                        }
                        else if (faultRequest is UpdateRelationshipRequest urr)
                        {
                            message =
                                $"Error while updating relationship {urr.Relationship.SchemaName}: {response.Fault.Message}";
                        }
                        else if (faultRequest is UpdateOptionSetRequest uosr)
                        {
                            message =
                                $"Error while updating optionset {uosr.OptionSet.Name}: {response.Fault.Message}";
                        }
                        else if (faultRequest is UpdateOptionValueRequest uovr)
                        {
                            if (!string.IsNullOrEmpty(uovr.OptionSetName))
                            {
                                message =
                                    $"Error while updating global optionset ({uovr.OptionSetName}) value ({uovr.Value}) label: {response.Fault.Message}";
                            }
                            else
                            {
                                message =
                                    $"Error while updating option ({uovr.Value}) label for attribute {uovr.AttributeLogicalName} ({uovr.EntityLogicalName}): {response.Fault.Message}";
                            }
                        }
                        else if (faultRequest is SetLocLabelsRequest sllr)
                        {
                            message =
                                $"Error while updating {sllr.AttributeName} of record {sllr.EntityMoniker.LogicalName} ({sllr.EntityMoniker.Id}): {response.Fault.Message}";
                        }
                        else
                        {
                            message = response.Fault.Message;
                        }

                        OnLog(new LogEventArgs(message) { Type = LogType.Error });
                    }
                }
            }
            else
            {
                e.SuccessCount += request.Requests.Count;
            }

            OnResult(e);

            InitMultipleRequest();
        }

        protected void InitMultipleRequest()
        {
            request = new ExecuteMultipleRequest
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = true,
                    ReturnResponses = false
                }
            };
        }
    }

    public class LogEventArgs : EventArgs
    {
        public LogEventArgs()
        {
        }

        public LogEventArgs(string message)
        {
            Message = message;
        }

        public string Message { get; set; }
        public LogType Type { get; set; }
    }

    public class TranslationProgressEventArgs : EventArgs
    {
        public int FailureCount { get; set; }
        public string SheetName { get; set; }
        public int SuccessCount { get; set; }
        public int TotalItems { get; set; }
    }

    public class TranslationResultEventArgs : EventArgs
    {
        public string Message { get; set; }
        public string SheetName { get; set; }
        public bool Success { get; set; }
    }
}