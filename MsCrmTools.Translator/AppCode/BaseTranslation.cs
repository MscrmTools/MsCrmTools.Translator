using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsCrmTools.Translator.AppCode
{
    public class TranslationResultEventArgs : EventArgs
    {
        public bool Success { get; set; }
        public string SheetName { get; set; }
        public string Message { get; set; }
    }
    public class BaseTranslation
    {
        public event EventHandler<TranslationResultEventArgs> Result;

        public virtual void OnResult(TranslationResultEventArgs e)
        {
            EventHandler<TranslationResultEventArgs> handler = Result;
            if (handler != null)
            {
                handler(this, e);
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(e.Message))
                {
                    File.AppendAllText("Logs\\ImportTranslations_" + DateTime.Now.Date.ToString("MMddyyyy") + ".log",
                        string.Format("{0}{1} - {2} - {3}", Environment.NewLine, e.SheetName, e.Success, e.Message));
                }
            }
            catch
            {
            }
        }
    }
}
