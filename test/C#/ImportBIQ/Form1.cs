using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Project = Microsoft.Office.Interop.MSProject;
using System.Diagnostics;
namespace ImportBIQ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.mpp";
            ofd.Filter = "Microsoft Project (*.mpp*)|*.mpp*";
            ofd.Title = "Выберите документ Project";

            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //Process.Start(ofd.FileName);
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void ButtShowOpenDialogBIQ_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.mpp";
            ofd.Filter = "Microsoft Project (*.mpp*)|*.mpp*";
            ofd.Title = "Выберите документ Project";

            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //Process.Start(ofd.FileName);
        }
    }
}
