using System;
using System.ComponentModel;
using System.IO;

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

            try
            {
                File.AppendAllText("Logs\\ImportTranslations_progress_" + DateTime.Now.Date.ToString("MMddyyyy") + ".log",
                      string.Format("{0}Progres - Overall:{1}, Item:{2}. Message:{3}", Environment.NewLine, pInfo.Overall, pInfo.Item, pInfo.Message));
            }
            catch { }
        }
    }
}
