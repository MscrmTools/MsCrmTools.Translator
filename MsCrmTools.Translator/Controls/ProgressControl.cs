using MsCrmTools.Translator.Properties;
using System.Windows.Forms;

namespace MsCrmTools.Translator.Controls
{
    public partial class ProgressControl : UserControl
    {
        public ProgressControl(string sheetName)
        {
            InitializeComponent();

            lblCount.Text = "0";
            lblError.Text = "0";
            lblSuccess.Text = "0";

            SheetName = sheetName;
            lblTitle.Text = sheetName;

            pictureBox4.Image = Resources.control_play_blue;
        }

        public int Count
        {
            set => lblCount.Text = value.ToString();
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
            pictureBox4.Image = succeeded ? Resources.tick : Resources.cancel;
        }
    }
}