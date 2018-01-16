using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharedPowerpointFavoritesPlugin.view
{
    public partial class InfoDialog : Form
    {

        public static readonly string INFO_TEXT = "This PowerPoint Addin was built in January 2018 by\nChristopher Rudoll <christopher@rudoll.net>\nfor\nPresentur.de";

        public InfoDialog()
        {
            InitializeComponent();
            this.Text = "About";
            var label = new Label();
            label.Text = INFO_TEXT;
            label.AutoSize = true;
            label.TextAlign = ContentAlignment.MiddleCenter;
            var actualSize = label.CreateGraphics().MeasureString(label.Text, label.Font);
            this.Width = (int) actualSize.Width + 20;
            this.Height = (int) actualSize.Height + 50;
            this.Controls.Add(label);

        }
    }
}
