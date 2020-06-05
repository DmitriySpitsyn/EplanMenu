using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    public partial class SelectKIT : Form
    {
        public SelectKIT(Form2 f)
        {
            InitializeComponent();
            Mainform = f;

        }
        public Form2 Mainform;
        public string selectNameKit { get; set; }
        public string selectLocation { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void SelectKIT_Shown(object sender, EventArgs e)
        {
           foreach(Form2.work f2w in Mainform.MainList)
            {
                dataGridView2.Rows.Add(f2w.NameKit);

            }
            foreach (Form2.part f2p in Mainform.location)
            {
                dataGridView3.Rows.Add(f2p.design,f2p.note);

            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            selectNameKit=dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString();
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            selectLocation = dataGridView3 [0, dataGridView3.CurrentRow.Index].Value.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            selectNameKit = "";
            selectLocation = "";
            Close();
        }
    }
}
