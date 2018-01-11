using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharedPowerpointFavoritesPlugin
{
    public partial class SharedFavView : Form
    {

        private static SharedFavView CURRENT_INSTANCE;

        public SharedFavView()
        {
            InitializeComponent();
        }

        private void SharedFavView_FormClosed(object sender, FormClosedEventArgs e)
        {
            CURRENT_INSTANCE = null;
        }

        private void SharedFavView_Shown(object sender, EventArgs e)
        {
            CURRENT_INSTANCE = this;
        }

        public static void ShowOrFocus()
        {
            if(CURRENT_INSTANCE == null)
            {
                new SharedFavView().Show();
            }
            else
            {
                CURRENT_INSTANCE.BringToFront();
            }
        }
    }
}
