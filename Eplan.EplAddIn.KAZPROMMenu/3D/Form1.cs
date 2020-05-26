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


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public struct functype
        {
            public string Page;
            public string Name;
            public string Designation;

        }
        public List<functype> dev = new List<functype>();
        public List<functype> devf = new List<functype>();


        string projectname = "";
        private void button1_Click(object sender, EventArgs e)
        {
            //Thread myThread = new Thread(new ParameterizedThreadStart(Update));
            Update(dev);
            /* myThread.Start(dev);
             myThread.Priority = ThreadPriority.Highest;
             myThread.Join();
             myThread.Abort();//прерываем поток
             myThread.Join(500);//таймаут на завершение*/
            updategui();



        }


        public void updategui()
        {


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


        void Update(object devin)
        {
            using (LockingStep oLS = new LockingStep())
            { // ... доступ к данным P8 ...

                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                projectname = CurrentProject.ProjectFullName;
                StorableObject[] storableObjects = Set.Selection;
                List<Page> Lpage = Set.GetSelectedPages().ToList();
                DMObjectsFinder dmObjectsFinder = new DMObjectsFinder(CurrentProject);
                List<Function3D> place3d;
                //place3d = dmObjectsFinder.GetFunctions3D(null).ToList();
                List<Function> place;
                //place = dmObjectsFinder.GetFunctions(null).ToList();
                List<Terminal> ltermlist;
                functype devl = new functype();
                ((List<functype>)devin).Clear();
                int id = 0;
                int id3d = 0;
                bool placefind = false;
                progressBar1.Value = 0;
                progressBar1.Maximum = Lpage.Count;

                for (int i = 0; i < Lpage.Count; i++)
                {
                    Type ltp = Lpage[i].GetType();
                    if (ltp.Name == "Page")
                    {
                        place = Lpage[i].Functions.ToList();
                        FunctionsFilter oterfilt = new FunctionsFilter();
                        oterfilt.Page = Lpage[i];
                        ltermlist = dmObjectsFinder.GetTerminals(oterfilt).ToList();
                        //MessageBox.Show(ltermlist.Count.ToString());
                        for (int j = 0; j < place.Count; j++)
                        {
                            id += 1;
                            if (place[j].IsMainFunction != true) { continue; } //// Проверяем является ли функция главной
                            //if (place[j].Properties.DESIGNATION_LOCATION.ToString() != listBox1.SelectedItem.ToString()) { continue; }
                            Functions3DFilter f3d = new Functions3DFilter();
                            Function3DPropertyList oFunction3DPropertyList = new Function3DPropertyList();
                            oFunction3DPropertyList.DESIGNATION_LOCATION = place[j].Properties.DESIGNATION_LOCATION;
                            f3d.SetFilteredPropertyList(oFunction3DPropertyList);
                            place3d = dmObjectsFinder.GetFunctions3D(f3d).ToList();
                            id3d += place3d.Count;
                            placefind = false;
                            for (int l = 0; l < place3d.Count; l++)
                            {
                                //MessageBox.Show("F"+ place[j].ObjectIdentifier.ToString()+"F3D "+ place3d[l].ObjectIdentifier+"\n"+ place[j].Name+" "+place3d[l].Name);

                                if (place[j].Properties.FUNC_IDENTDEVICETAGWITHOUTSTRUCTURES == place3d[l].Properties.FUNC_IDENTDEVICETAGWITHOUTSTRUCTURES)
                                {
                                    placefind = true;
                                    break;
                                }
                            }
                            if (placefind == false)
                            {
                                devl.Name = place[j].Name;
                                devl.Page = Lpage[i].Name;
                                devl.Designation = place[j].Properties.DESIGNATION_LOCATION;
                                ((List<functype>)devin).Add(devl);
                            }


                        }

                        for (int t = 0; t < ltermlist.Count; t++)
                        {
                            if (ltermlist[t].IsMainTerminal != true) { continue; } //// Проверяем является ли функция главной
                            //if (place[j].Properties.DESIGNATION_LOCATION.ToString() != listBox1.SelectedItem.ToString()) { continue; }
                            Functions3DFilter f3d = new Functions3DFilter();
                            Function3DPropertyList oFunction3DPropertyList = new Function3DPropertyList();
                            oFunction3DPropertyList.DESIGNATION_LOCATION = ltermlist[t].Properties.DESIGNATION_LOCATION;
                            f3d.SetFilteredPropertyList(oFunction3DPropertyList);
                            place3d = dmObjectsFinder.GetFunctions3D(f3d).ToList();
                            id3d += place3d.Count;
                            // MessageBox.Show(ltermlist[t].Name.ToString());
                            placefind = false;
                            for (int l = 0; l < place3d.Count; l++)
                            {

                                if (ltermlist[t].Name.ToString() == place3d[l].Name.ToString())
                                {
                                    placefind = true;
                                    break;
                                }
                            }
                            if (placefind == false)
                            {
                                devl.Name = ltermlist[t].Name;
                                devl.Page = Lpage[i].Name;
                                devl.Designation = ltermlist[t].Properties.DESIGNATION_LOCATION;
                                ((List<functype>)devin).Add(devl);
                            }

                        }

                    }
                    progressBar1.Value += 1;
                }
                label4.Text = id.ToString();
                label7.Text = id3d.ToString();
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
                    listBox2.Items.Add("Страница №: " + dev[i].Page + " " + dev[i].Name);
                }
                else
                {
                    if (listBox1.SelectedItem.ToString() == dev[i].Designation)
                    {
                        // MessageBox.Show(listBox1.SelectedItem.ToString() +"---"+ dev[i].Designation);
                        devf.Add(dev[i]);
                        listBox2.Items.Add("Страница №: " + dev[i].Page + " " + dev[i].Name);
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






