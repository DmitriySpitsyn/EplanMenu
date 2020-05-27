using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Eplan.EplApi.DataModel.EObjects;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public class part
        {
            public string partnr { get; set; }
            public int pcount { get; set; }
            public PropertyValue design { get; set; }
            public PropertyValue note { get; set; }
        }
        public List<part> filtpart = new List<part>();
        public List<part> filtpart2 = new List<part>();
        public List <DeviceListEntry> Devlist=new List<DeviceListEntry>() ;
        BackgroundWorker bw;
        private void button1_Click(object sender, EventArgs e)
        {
            bw = new BackgroundWorker();
            bw.DoWork += (obj, ea)=>partup(1);
            bw.RunWorkerAsync();

           
        }

       private async void partup(int times)
        {
            using (LockingStep oLS = new LockingStep())
            { // ... доступ к данным P8 ...

                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                StorableObject[] storableObjects = Set.Selection;
                List<Page> Lpage = Set.GetSelectedPages().ToList();
                List<Function> func = new List<Function>();
                List<Terminal> term = new List<Terminal>();
                FunctionsFilter oterfilt = new FunctionsFilter();

                List<ArticleReference> articref = new List<ArticleReference>();
                DMObjectsFinder dmObjectsFinder = new DMObjectsFinder(CurrentProject);
                //List<Eplan.EplApi.DataModel.EObjects.PLC> PLCs = new List<Eplan.EplApi.DataModel.EObjects.PLC>();
                //FunctionsFilter ofuncfilter = new FunctionsFilter();
                filtpart.Clear();
                bool searchPLC = false;
                progressBar1.Maximum = Lpage.Count - 1;
                for (int p = 0; p < Lpage.Count; p++)
                {
                    progressBar1.Value = p;
                    func = Lpage[p].Functions.ToList();
                    foreach (Function f in func)
                    {
                        if (f.IsMainFunction != true) { continue; }
                        articref = f.ArticleReferences.ToList();
                        foreach (ArticleReference ar in articref)
                        {

                            searchPLC = false;

                            foreach (part cpart in filtpart)
                            {
                                if ((cpart.partnr == ar.PartNr) & (f.Properties.DESIGNATION_FULLLOCATION == cpart.design))
                                {
                                    searchPLC = true;

                                    cpart.pcount += ar.Properties.ARTICLEREF_COUNT;
                                    break;
                                }

                            }
                            if (searchPLC == false)
                            {


                                filtpart.Add(new part() { partnr = ar.PartNr, pcount = ar.Properties.ARTICLEREF_COUNT, design = f.Properties.DESIGNATION_FULLLOCATION });
                            }

                        }
                    }
                    oterfilt.Page = Lpage[p];
                    term = dmObjectsFinder.GetTerminals(oterfilt).ToList();
                    foreach (Terminal ft in term)
                    {
                        if (ft.IsMainTerminal != true) { continue; }
                        articref = ft.ArticleReferences.ToList();
                        foreach (ArticleReference art in articref)
                        {
                            searchPLC = false;
                            foreach (part cpart in filtpart)
                            {
                                if ((cpart.partnr == art.PartNr) & (ft.Properties.DESIGNATION_FULLLOCATION == cpart.design))
                                {
                                    searchPLC = true;

                                    cpart.pcount += art.Properties.ARTICLEREF_COUNT;
                                    break;
                                }
                            }
                            if (searchPLC == false)
                            {

                                filtpart.Add(new part() { partnr = art.PartNr, pcount = art.Properties.ARTICLEREF_COUNT, design = ft.Properties.DESIGNATION_FULLLOCATION });
                            }
                        }
                    }

                }

                /* for (int i=0;i<filtpart.Count;i++)
                 {
                     listBox1.Items.Add(i.ToString()+" "+filtpart[i].partnr + "  " + filtpart[i].pcount.ToString() + "***" + filtpart[i].design);

                 }*/
          

            }

        }
        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

       

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (LockingStep oLS = new LockingStep())
            { // ... доступ к данным P8 ...

                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                StorableObject[] storableObjects = Set.Selection;
                foreach(StorableObject f in storableObjects)
                {
                    MessageBox.Show(f.GetType().ToString());

                }
            }
            }

        private void button3_Click(object sender, EventArgs e)
        {
            using (LockingStep oLS = new LockingStep())
            { // ... доступ к данным P8 ...

               // dataGridView1.Rows.Clear();
                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                DeviceService devservice = new DeviceService();
                Devlist = devservice.GetAllDeviceListItems(CurrentProject).ToList();
                int j = 0;
                dataGridView2.Rows.Clear();
                dataGridView2.Rows.Add(1, "Общий список изделий", "---");
                bool searchgesign = false;
                progressBar1.Maximum = Devlist.Count;
                foreach(DeviceListEntry f in Devlist)
                {
                    progressBar1.Value += 1;
                    searchgesign = false;
                    for (int i=0;i< dataGridView2.Rows.Count; i++)
                    {
                        //MessageBox.Show(dataGridView2[0, i].Value.ToString()+" "+ dataGridView2[1, i].Value.ToString()+" " + dataGridView2[2, i].Value.ToString());
                        if ((f.Properties.DEVICELISTENTRY_PLANT==dataGridView2[1,i].Value .ToString())& (dataGridView2[1, i].Value!=null))
                        {
                            searchgesign = true;
                            break;
                        }
                    }
                    if (searchgesign == false)
                    {
                        dataGridView2.Rows.Add(dataGridView2.Rows.Count, f.Properties.DEVICELISTENTRY_PLANT, "--");
                        //listBox1.Items.Add(f.Properties.DEVICELISTENTRY_PLANT);
                    }
                   // dataGridView1.Rows.Add(j, f.Properties.DEVICELISTENTRY_PARTNR, f.Properties.DEVICELISTENTRY_COUNTALLOWED);
                }
            }

        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            filtpart2.Clear();
            dataGridView1.Rows.Clear();
            bool desloctrue = false;
            foreach (DeviceListEntry f1 in Devlist)
            {
                if (dataGridView2.CurrentRow.Index == 0)
                {
                    desloctrue = false;
                    foreach (part f2 in filtpart2)
                    {
                        if (f1.Properties.DEVICELISTENTRY_PARTNR == f2.partnr)
                        {
                            f2.pcount += f1.Properties.DEVICELISTENTRY_COUNTALLOWED;
                            desloctrue = true;
                            break;
                        }
                    }
                    if (desloctrue == false)
                    {
                        filtpart2.Add(new part() { partnr = f1.Properties.DEVICELISTENTRY_PARTNR, pcount = f1.Properties.DEVICELISTENTRY_COUNTALLOWED });
                    }
                }
                else
                {

                    if (dataGridView2.CurrentCell.Value.ToString() == f1.Properties.DEVICELISTENTRY_PLANT)
                    {
                        desloctrue = false;
                        foreach (part f2 in filtpart2)
                        {
                            if (f1.Properties.DEVICELISTENTRY_PARTNR == f2.partnr)
                            {
                                f2.pcount += f1.Properties.DEVICELISTENTRY_COUNTALLOWED;
                                desloctrue = true;
                                break;
                            }
                        }
                        if (desloctrue == false)
                        {
                            filtpart2.Add(new part() { partnr = f1.Properties.DEVICELISTENTRY_PARTNR, pcount = f1.Properties.DEVICELISTENTRY_COUNTALLOWED });
                        }

                    }
                }

            }
            for (int j = 0; j < filtpart2.Count; j++)
            {
                dataGridView1.Rows.Add(j + 1, filtpart2[j].partnr, filtpart2[j].pcount);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
