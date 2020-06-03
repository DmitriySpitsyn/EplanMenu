using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Eplan.EplApi.DataModel.EObjects;
using Eplan.EplApi.Base;

namespace Eplan.EplAddIn.KAZPROMMenu
{


    public partial class Form1ver2 : Form
    {
        public Form1ver2()
        {
            InitializeComponent();
        }
        public struct functype
        {
            public string Page;
            public string Name;
            public string Designation;
            public bool notfull;

        }
        public class part
        {
            public string partnr { get; set; }
            public int pcount { get; set; }
            public PropertyValue design { get; set; }
        }
        public List<functype> dev = new List<functype>();
        public List<functype> devf = new List<functype>();
        public int countelement=0;

        string projectname = "";
        BackgroundWorker bw;
        private void button1_Click(object sender, EventArgs e)
        {

            bw = new BackgroundWorker();
            if (bw.IsBusy) return;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.RunWorkerAsync();
            button1.Enabled = false;

        }
        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            using (LockingStep oLS = new LockingStep())
            {  // ... доступ к данным P8 ...

                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                projectname = CurrentProject.ProjectFullName;
                StorableObject[] storableObjects = Set.Selection;
                List<Page> Lpage = Set.GetSelectedPages().ToList();
                List<Function> func = new List<Function>();
                List<Terminal> term = new List<Terminal>();
                FunctionsFilter oterfilt = new FunctionsFilter();
                functype devl = new functype();
                List<ArticleReference> articref = new List<ArticleReference>();
                DMObjectsFinder dmObjectsFinder = new DMObjectsFinder(CurrentProject);
                //List<Eplan.EplApi.DataModel.EObjects.PLC> PLCs = new List<Eplan.EplApi.DataModel.EObjects.PLC>();
                //FunctionsFilter ofuncfilter = new FunctionsFilter();
                dev.Clear();
                bool searchPLC = false;
                progressBar1.Maximum = Lpage.Count - 1;
                int arcount3d = 0;
                for (int p = 0; p < Lpage.Count; p++)
                {
                    progressBar1.Value = p;
                    func = Lpage[p].Functions.ToList();
                    foreach (Function f in func)
                    {
                        if (checkBox1.Checked == true)
                        {
                            if (f.IsMainFunction != true) { continue; }
                        }
                        else
                        {
                            if ((f.IsMainFunction != true) || (f.Properties.FUNC_CATEGORY_GROUP_ID == "400/1/1")) { continue; }
                        }

                        countelement += 1;
                        articref = f.ArticleReferences.ToList();
                        searchPLC = false;
                        arcount3d = 0;
                        foreach (ArticleReference ar in articref)
                        {
                            if (ar.Properties.ARTICLEREF_COUNT_NOTPLACED_3D != 0)
                            {
                                searchPLC = true;
                                if (ar.Properties.ARTICLEREF_COUNT - ar.Properties.ARTICLEREF_COUNT_NOTPLACED_3D == 0)
                                {
                                    arcount3d += 1;
                                }
                            }
                        }

                        if ((searchPLC == true) || (articref.Count == 0))
                        {
                            devl.Name = f.Name;
                            devl.Page = Lpage[p].Name;
                            devl.Designation = f.Properties.DESIGNATION_FULLLOCATION;
                            if (arcount3d == articref.Count) { devl.notfull = false; } else { devl.notfull = true; }
                            dev.Add(devl);
                        }
                    }


                    oterfilt.Page = Lpage[p];
                    term = dmObjectsFinder.GetTerminals(oterfilt).ToList();
                    foreach (Terminal ft in term)
                    {
                        if (ft.IsMainTerminal != true) { continue; }
                        countelement += 1;
                        articref = ft.ArticleReferences.ToList();
                        searchPLC = false;
                        arcount3d = 0;
                        foreach (ArticleReference ar in articref)
                        {
                            if (ar.Properties.ARTICLEREF_COUNT_NOTPLACED_3D != 0)
                            {
                                searchPLC = true;
                                if (ar.Properties.ARTICLEREF_COUNT - ar.Properties.ARTICLEREF_COUNT_NOTPLACED_3D == 0)
                                {
                                    arcount3d += 1;
                                }
                            }

                        }

                        if ((searchPLC == true) || (articref.Count == 0))
                        {
                            devl.Name = ft.Name;
                            devl.Page = Lpage[p].Name;
                            devl.Designation = ft.Properties.DESIGNATION_FULLLOCATION;
                            if (arcount3d == articref.Count) { devl.notfull = false; } else { devl.notfull = true; }
                            dev.Add(devl);
                        }
                    }
                }

                updategui();
                button1.Enabled = true;
            }

        }


        public void updategui()
        {

            label4.Text = countelement.ToString();
            listBox1.Items.Clear();
            listBox1.Items.Add("Все_");
            for (int i = 0; i < dev.Count; i++)
            {
                bool desloctrue = false;
                for (int c = 0; c < listBox1.Items.Count; c++)
                {
                    if (listBox1.Items[c].ToString() == dev[i].Designation)
                    {
                        desloctrue = true;
                    }
                }
                if (desloctrue != true & (dev[i].Designation != ""))
                {
                    listBox1.Items.Add(dev[i].Designation);
                }
            }
        }


 


        private void button2_Click(object sender, EventArgs e)
        {



        }



        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'engineeringDataSet.Lazarus' table. You can move, or remove it, as needed.


        }





        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            devf.Clear();
            for (int i = 0; i < dev.Count; i++)
            {
                if (listBox1.SelectedIndex == 0)
                {

                    devf.Add(dev[i]);
                    if (dev[i].notfull==false)
                    { listBox2.Items.Add("Страница №: " + dev[i].Page + " " + dev[i].Name); }
                    else
                    { listBox2.Items.Add("Страница №: " + dev[i].Page + " " + dev[i].Name+"  НЕ ПОЛНОСТЬЮ!"); }

                }
                else
                {
                    if (listBox1.SelectedItem.ToString() == dev[i].Designation)
                    {
                        // MessageBox.Show(listBox1.SelectedItem.ToString() +"---"+ dev[i].Designation);
                        devf.Add(dev[i]);
                        if (dev[i].notfull == false)
                        { listBox2.Items.Add("Страница №: " + dev[i].Page + " " + dev[i].Name); }
                        else
                        { listBox2.Items.Add("Страница №: " + dev[i].Page + " " + dev[i].Name + "  НЕ ПОЛНОСТЬЮ!"); }
                    }
                }
            }
            label3.Text = listBox2.Items.Count.ToString();



        }

        private void listBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Edit oedit = new Edit();
            using (LockingStep oLS = new LockingStep())
            {
                if (projectname != "")
                {
                    //oedit.OpenPageWithName(projectname, devf[listBox2.SelectedIndex].Page);

                    oedit.OpenPageWithNameAndFunctionName(projectname, devf[listBox2.SelectedIndex].Page, devf[listBox2.SelectedIndex].Name);
                    oedit.SetFocusToGED();
                }

            }

        }



        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_3(object sender, EventArgs e)
        {
            
        }

    }
}






