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
        }
        public List<part> filtpart = new List<part>();
        public List<part> filtpart2 = new List<part>();
        private void button1_Click(object sender, EventArgs e)
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
                
                List < ArticleReference > articref= new List<ArticleReference>();
                DMObjectsFinder dmObjectsFinder = new DMObjectsFinder(CurrentProject);
                //List<Eplan.EplApi.DataModel.EObjects.PLC> PLCs = new List<Eplan.EplApi.DataModel.EObjects.PLC>();
                //FunctionsFilter ofuncfilter = new FunctionsFilter();
                filtpart.Clear();
                bool searchPLC = false;
                progressBar1.Maximum = Lpage.Count-1;
                for (int p=0;p<Lpage.Count;p++)
                {
                    progressBar1.Value = p;
                    func = Lpage[p].Functions.ToList();
                    foreach(Function f in func)
                    {
                        if (f.IsMainFunction != true) { continue; }
                        articref = f.ArticleReferences.ToList();
                        foreach(ArticleReference ar in articref)
                        {

                            searchPLC = false;
                            
                            foreach (part cpart in filtpart)
                            {
                                if ((cpart.partnr == ar.PartNr) &(f.Properties.DESIGNATION_LOCATION == cpart.design))
                                   {
                                    searchPLC = true;
                                    
                                    cpart.pcount += ar.Properties.ARTICLEREF_COUNT;
                                        break;
                                    }

                            }
                            if (searchPLC == false)
                            {


                                filtpart.Add(new part() { partnr = ar.PartNr, pcount = ar.Properties.ARTICLEREF_COUNT, design = f.Properties.DESIGNATION_LOCATION });
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
                                if ((cpart.partnr == art.PartNr) & (ft.Properties.DESIGNATION_LOCATION == cpart.design))
                                {
                                    searchPLC = true;
                                    
                                    cpart.pcount += art.Properties.ARTICLEREF_COUNT;
                                    break;
                                }
                            }
                            if (searchPLC == false)
                            {

                                filtpart.Add(new part() { partnr = art.PartNr, pcount = art.Properties.ARTICLEREF_COUNT, design = ft.Properties.DESIGNATION_LOCATION });
                            }
                        }
                    }
                    
                }
               
               /* for (int i=0;i<filtpart.Count;i++)
                {
                    listBox1.Items.Add(i.ToString()+" "+filtpart[i].partnr + "  " + filtpart[i].pcount.ToString() + "***" + filtpart[i].design);

                }*/
                listBox1.Items.Clear();
                listBox1.Items.Add("Все_");
                for (int i = 0; i < filtpart.Count; i++)
                {
                    bool desloctrue = false;
                    for (int c = 0; c < listBox1.Items.Count; c++)
                    {
                        if (listBox1.Items[c].ToString() == filtpart[i].design)
                        {
                            desloctrue = true;
                        }
                    }
                    if (desloctrue != true & (filtpart[i].design != ""))
                    {
                        listBox1.Items.Add(filtpart[i].design);
                    }
                }



            }
        }

       
        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            filtpart2.Clear();
            bool desloctrue = false;
            foreach(part f1 in filtpart)
            {
                if (listBox1.SelectedIndex == 0)
                {
                    desloctrue = false;
                    foreach (part f2 in filtpart2)
                    {
                        if (f1.partnr == f2.partnr)
                        {
                            f2.pcount += f1.pcount;
                            desloctrue = true;
                            break;
                        }
                    }
                    if (desloctrue==false)
                    {
                        filtpart2.Add(new part() { partnr = f1.partnr, pcount = f1.pcount});
                    }
                   }
                else
                {

                    if (listBox1.SelectedItem.ToString() == f1.design)
                    {
                        desloctrue = false;
                        foreach (part f2 in filtpart2)
                        {
                            if (f1.partnr == f2.partnr)
                            {
                                f2.pcount += f1.pcount;
                                desloctrue = true;
                                break;
                            }
                        }
                        if (desloctrue == false)
                        {
                            filtpart2.Add(new part() { partnr = f1.partnr, pcount = f1.pcount });
                        }

                    }
                }
            }
            foreach(part f2 in filtpart2)
            {
                listBox2.Items.Add(f2.partnr + "       " + f2.pcount);
            }
        }
    }
}
