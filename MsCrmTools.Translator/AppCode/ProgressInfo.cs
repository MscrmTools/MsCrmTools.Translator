using System.ComponentModel;

namespace MsCrmTools.Translator.AppCode
{
    public class ProgressInfo
    {
        public int Overall { get; set; }

        public int Item { get; set; }

        public string Message { get; set; }
    }

    public static class Extensions
    {
        public static void ReportProgressIfPossible(this BackgroundWorker worker, int progress, ProgressInfo pInfo)
        {
            if (worker != null && worker.WorkerReportsProgress)
            {
                worker.ReportProgress(progress, pInfo);
            }
        }
    }
}
