using MsCrmTools.Translator.Properties;
using System.Windows.Forms;

namespace MsCrmTools.Translator.Controls
{
    public partial class ProgressControl : UserControl
    {
        private ToolTip tooltip = new ToolTip();

        public ProgressControl(string sheetName)
        {
            InitializeComponent();

            lblCount.Text = "0";
            lblError.Text = "0";
            lblSuccess.Text = "0";

            SheetName = sheetName;
            lblTitle.Text = sheetName;

            pbProgress.Image = Resources.control_play_blue;
        }

        public int Count
        {
            set
            {
                lblCount.Text = value.ToString();
                tooltip.SetToolTip(lblCount, $"Number of labels to be translated : {lblCount.Text}\n\n(Might be different from the number of attributes, views, etc. Attributes, optionset values, etc. have label and description)");
                tooltip.SetToolTip(pbTotal, $"Number of labels to be translated : {lblCount.Text}\n\n(Might be different from the number of attributes, views, etc. Attributes, optionset values, etc. have label and description)");
            }
        }

        public int Error
        {
            set => lblError.Text = value.ToString();
            get => int.Parse(lblError.Text);
        }

        public string SheetName { get; }

        public int Success
        {
            set => lblSuccess.Text = value.ToString();
        }

        public void End(bool succeeded)
        {
            pbProgress.Image = succeeded ? Resources.tick : Resources.cancel;
        }
    }
}