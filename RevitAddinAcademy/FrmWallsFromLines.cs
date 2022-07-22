using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAddinAcademy
{
    public partial class FrmWallsFromLines : Form
    {
        public FrmWallsFromLines(List<string> wallTypes, List<string> lineStyles)
        {
            InitializeComponent();

            foreach(string wallType in wallTypes)
            {
                this.cmbWallTypes.Items.Add(wallType);
            }

            foreach(string lineStyle in lineStyles)
            {
                this.cmbLineStyles.Items.Add(lineStyle);
            }

            this.cmbWallTypes.SelectedIndex = 0;
            this.cmbLineStyles.SelectedIndex = 0;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public string GetSelectedWallType()
        {
            return cmbWallTypes.SelectedItem.ToString();
        }

        public string GetSelectedLineStyle()
        {
            return cmbLineStyles.SelectedItem.ToString();
        }

        public double GetWalLHeight()
        {
            double returnValue;

            if(double.TryParse(tbxWallHeight.Text, out returnValue) == true)
            {
                return returnValue;
            }

            return 20;
        }

        public bool AreWallsStructural()
        {
            return cbxStructural.Checked;
        }
    }
}
