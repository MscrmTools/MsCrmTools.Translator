using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace MsCrmTools.Translator.AppCode
{
    public static class Extensions
    {
        public static List<Guid> GetSolutionComponentObjectIds(this IOrganizationService service, Guid solutionId, int type)
        {
            return service.RetrieveMultiple(new QueryExpression("solutioncomponent")
            {
                ColumnSet = new ColumnSet("objectid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("componenttype", ConditionOperator.Equal, type),
                        new ConditionExpression("solutionid", ConditionOperator.Equal, solutionId)
                    }
                }
            }).Entities.Select(e => e.GetAttributeValue<Guid>("objectid")).ToList();
        }

        public static void ReportProgressIfPossible(this BackgroundWorker worker, int progress, ProgressInfo pInfo)
        {
            if (worker != null && worker.WorkerReportsProgress)
            {
                worker.ReportProgress(progress, pInfo);
            }
        }
    }
}