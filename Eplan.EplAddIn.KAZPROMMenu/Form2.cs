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
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.MasterData;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.InteropServices;

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
            public string partnr { get; set; }//Заказной номер
            public int pcount { get; set; } //Количество
            public int spare { get; set; } //Резерв
            public string design { get; set; }//Название местоположения
            public string descrip { get; set; } //описание изделия
            public string note { get; set; }// описание местоположения
            public string erpn { get; set; }// номер ERP.Артикул в 1С
            public int price1 { get; set; }// цена в 1 валюте(тенге)
            public int price2 { get; set; }// цена в 2 валюте(евро)
        }
        public class work{
            public int id { get; set; }//Номер позиции
            public string NameKit { get; set; }//Название комплекта
            public int pcount { get; set; } //Количество 
            public List<part> parts { get; set; } //Список запчастей в комплекте
        }
        public class SpecPrice
        {
            public int id { get; set; }//Номер позиции
            public string NameKit { get; set; }//Название комплекта
            public int pcount { get; set; } //Количество 
            public long Price1 { get; set; } //Цена валюта1  
            public long Price2 { get; set; } //Цена валюта2
            public long COSTPrice1 { get; set; } //Стоимость валюта1 
            public long COSTPrice2 { get; set; } //Стоимость валюта2    
        }
        public List<work> MainList= new List<work>();
        public List<part> filtpart = new List<part>();
        public List<part> filtpart2 = new List<part>();
       // public List<part> filtdevice = new List<part>();
        public List<part> filtdevice2 = new List<part>();
        public List<SpecPrice> SpecALL = new List<SpecPrice>();
        public List <DeviceListEntry> Devlist=new List<DeviceListEntry>() ;
        public bool blockupdatelist=false;
        BackgroundWorker bw ;
        public Project CurProj;
        private void button1_Click(object sender, EventArgs e)
        {
            


        }


        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            using (LockingStep oLS = new LockingStep())
            { // ... доступ к данным P8 ...
                button1.Enabled = false;
                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                CurProj= Set.GetCurrentProject(true);
                //StorableObject[] storableObjects = Set.Selection;
                List<Page> Lpage = Set.GetSelectedPages().ToList();
                List<Function> func = new List<Function>();
                List<Function3D> func3d = new List<Function3D>();
                List<Terminal> term = new List<Terminal>();
                FunctionsFilter oterfilt = new FunctionsFilter();
                MultiLangString mlstring = new MultiLangString();
                List<ArticleReference> articref = new List<ArticleReference>();
                DMObjectsFinder dmObjectsFinder = new DMObjectsFinder(CurrentProject);
                Functions3DFilter f3dfilter = new Functions3DFilter();
                
                func3d = dmObjectsFinder.GetFunctions3D(f3dfilter).ToList();
                //List<Eplan.EplApi.DataModel.EObjects.PLC> PLCs = new List<Eplan.EplApi.DataModel.EObjects.PLC>();
                //FunctionsFilter ofuncfilter = new FunctionsFilter();
                filtpart.Clear();
                bool searchPLC = false;
                progressBar1.Maximum = Lpage.Count - 1;
                blockupdatelist = true;
                ArticleReferencesFilter arfilter = new ArticleReferencesFilter();
                articref = dmObjectsFinder.GetArticleReferences(arfilter).ToList();
                string erpbuf;
                int price1;
                int price2;
                string bufdecp = "";
                progressBar1.Maximum = articref.Count;
                progressBar1.Value = 0;
                             
                DeviceService devservice = new DeviceService();
                Devlist = devservice.GetAllDeviceListItems(CurrentProject).ToList();
                MainList.Clear();
                dataGridView2.Rows.Clear();
                MainList.Add(new work()
                {
                    id = 0,
                    NameKit = "Общий список изделий",
                    pcount =0,
                    parts = new List<part>()
                });
                bool searchgesign = false;
                //progressBar1.Maximum = Devlist.Count+1;
                string buf = "";
                blockupdatelist = true;
                #region KITFound// Ищем Комплекты
                {
                    foreach (DeviceListEntry f in Devlist)
                    {
                        try
                        {
                            if (f.Properties.DEVICELISTENTRY_PARTNR.ToString() == "") { continue; }
                        }
                        catch (EmptyPropertyException)
                        {
                            continue;

                        }


                        searchgesign = false;
                        buf = "";

                        for (int s = 0; s < f.Properties.DEVICELISTENTRY_PARTNR.ToString().Length; s++)
                        {
                            if (f.Properties.DEVICELISTENTRY_PARTNR.ToString()[s] == '.')

                            {
                                break;
                            }
                            else
                            {
                                buf += f.Properties.DEVICELISTENTRY_PARTNR.ToString()[s];
                            }
                        }

                        if (buf == "KIT")
                        {
                            //filtdevice.Add(new part() { partnr = f.Properties.DEVICELISTENTRY_PARTNR.ToString().Remove(0, 4), pcount = f.Properties.DEVICELISTENTRY_COUNTALLOWED });
                            MainList.Add(new work()
                            {
                                id = 0,
                                NameKit = f.Properties.DEVICELISTENTRY_PARTNR.ToString().Remove(0, 4),
                                pcount = f.Properties.DEVICELISTENTRY_COUNTALLOWED,
                                parts = new List<part>()
                            });
                            continue;
                        }

                    }

                }
                #endregion
                #region FoundDeviceforKIT// Заполняем комплектность
                {
                    foreach (DeviceListEntry f1 in Devlist)
                    {
                        if (MainList.Count == 0) { break; }
                        try
                        {
                            if (f1.Properties.DEVICELISTENTRY_PARTNR.ToString() == "") { continue; }
                        }
                        catch (EmptyPropertyException)
                        {
                            continue;
                        }
                        buf = "";
                        for (int s = 0; s < f1.Properties.DEVICELISTENTRY_PARTNR.ToString().Length; s++)
                        {
                            if (f1.Properties.DEVICELISTENTRY_PARTNR.ToString()[s] == '.')

                            {
                                break;
                            }
                            else
                            {
                                buf += f1.Properties.DEVICELISTENTRY_PARTNR.ToString()[s];
                            }
                        }

                        //MessageBox.Show(buf + "---" + buf.Length.ToString() + "---" + f1.Properties.DEVICELISTENTRY_PARTNR.ToString());
                        if (buf == "KIT")
                        {
                            continue;
                        }
                        bool desloctrue = false;
                        int k = 1;// коэфициент умножения на комплекты
                                  //MessageBox.Show(filtpart.Count.ToString() + "---" + f1.Properties.DEVICELISTENTRY_PLANT);
                        foreach (work KIT in MainList)
                        {
                            // MessageBox.Show("--" + KIT.partnr + "--" + k.ToString() + "--" + f1.Properties.DEVICELISTENTRY_PLANT + "--" + f1.Properties.DEVICELISTENTRY_PARTNR);

                            if (KIT.NameKit == f1.Properties.DEVICELISTENTRY_PLANT)
                            {
                                desloctrue = false;
                                foreach (part f2 in KIT.parts)
                                {

                                    if (f1.Properties.DEVICELISTENTRY_PARTNR == f2.partnr)
                                    {
                                        f2.pcount += f1.Properties.DEVICELISTENTRY_COUNTALLOWED ;
                                        desloctrue = true;
                                        break;
                                    }
                                }
                                if (desloctrue == false)
                                {
                                    try
                                    {
                                        bufdecp = f1.Properties.DEVICELISTENTRY_DESCRIPTION;
                                    }
                                    catch (EmptyPropertyException)
                                    {

                                    }
                                    KIT.parts.Add(new part() { partnr = f1.Properties.DEVICELISTENTRY_PARTNR, pcount = f1.Properties.DEVICELISTENTRY_COUNTALLOWED , descrip = bufdecp });
                                }
                                k = KIT.pcount;
                                break;
                            }
                        }
                            foreach (part f2 in MainList[0].parts)
                            {
                            desloctrue = false;
                            if (f1.Properties.DEVICELISTENTRY_PARTNR == f2.partnr)
                                {
                                    f2.pcount += (f1.Properties.DEVICELISTENTRY_COUNTALLOWED * k);
                                    desloctrue = true;
                                    break;
                                }
                            }
                            if (desloctrue == false)
                            {
                                try
                                {
                                    bufdecp = f1.Properties.DEVICELISTENTRY_DESCRIPTION;
                                }
                                catch (EmptyPropertyException)
                                {

                                }
                                MainList[0].parts.Add(new part() { partnr = f1.Properties.DEVICELISTENTRY_PARTNR, pcount = f1.Properties.DEVICELISTENTRY_COUNTALLOWED * k, descrip = bufdecp });
                            }
                        
                    }
                    }
                #endregion
                
                    for (int j = 0; j < MainList.Count; j++)
                {
                    MainList[j].id = j;
                    dataGridView2.Rows.Add(MainList[j].id, MainList[j].NameKit, MainList[j].pcount);
                    if (j == 0) continue;
                    SpecALL.Add(new SpecPrice() {
                        id = MainList[j].id,
                        NameKit = MainList[j].NameKit,
                        pcount = MainList[j].pcount,
                        Price1 = 0,
                        Price2 = 0,
                        COSTPrice1 = 0,
                        COSTPrice2 = 0 });
                }
                #region //Сравниваю с проектом
                {
                    foreach (ArticleReference ar in articref)
                    {
                        //List<StorableObject> storableObjects = ar.CrossReferencedObjectsAll.ToList();
                        //StorableObject[] storableObjects =ar.CrossReferencedObjectsAll
                        searchPLC = false;
                        //listBox1.Items.Add(ar.PartNr + "  " + ar.Properties.ARTICLEREF_COUNT);


                        foreach (part cpart in filtpart)
                        {
                            
                            if ((cpart.partnr == ar.PartNr) & (ar.Properties.DESIGNATION_FULLLOCATION == cpart.design))
                            {
                                searchPLC = true;
                                cpart.pcount += ar.Properties.ARTICLEREF_COUNT;
                                break;
                            }

                        }
                        if (searchPLC == false)
                        {

                            mlstring = ar.Properties.DESIGNATION_FULLLOCATION_DESCR.ToMultiLangString();
                           /* try
                            {
                                erpbuf = ar.Properties.ARTICLE_ERPNR;
                            }
                            catch (EmptyPropertyException)
                            {
                                erpbuf = "";

                            }
                            try
                            {
                                price1 = ar.Properties.ARTICLE_SALESPRICE_1;
                            }
                            catch (EmptyPropertyException)
                            {
                                price1 = 0;

                            }
                            try
                            {
                                price2 = ar.Properties.ARTICLE_SALESPRICE_2;
                            }
                            catch (EmptyPropertyException)
                            {
                                price2 = 0;

                            }*/
                            filtpart.Add(new part()
                            {
                                partnr = ar.PartNr,
                                pcount = ar.Properties.ARTICLEREF_COUNT,
                                design = ar.Properties.DESIGNATION_FULLLOCATION,
                                note = mlstring.GetStringToDisplay(ISOCode.Language.L_ru_RU),
                                /*erpn = erpbuf,
                                price1 = price1,
                                price2 = price2,*/
                            });




                        }



                        progressBar1.Value++;
                    }

                    // Подсчет резерва 
                    string bufname1 = "";
                    string bufname2 = "";
                    foreach (work KIT in MainList)
                    {
                        foreach (part rowkit in KIT.parts)
                        {
                            rowkit.spare = rowkit.pcount;
                            bufname1 = rowkit.partnr;
                            bufname1 = bufname1.Replace(" ", "");
                            foreach (part cpart in filtpart)
                            {
                                bufname2 = cpart.partnr;
                                bufname2 = bufname2.Replace(" ", "");
                                if (bufname1 == bufname2)
                                {
                                    rowkit.spare -= cpart.pcount;
                                }
                            }                         
                        }
                    }
                   


                }
                #endregion
                #region // Расставление ERP и формирование цен
                {
                    progressBar1.Maximum += MainList[0].parts.Count;
                    MDPartsManagement pm = new MDPartsManagement();
                    MDPartsDatabase db = pm.OpenDatabase();
                    MDPart mdpart;
                    //progressBar1.Value = 0;
                    foreach (part rowkit in MainList[0].parts)
                    {
                        mdpart = db.GetPart(rowkit.partnr);
                        progressBar1.Value++;
                        try
                        {
                            //rowkit.spare -= cpart.pcount;
                            rowkit.erpn = mdpart.Properties.ARTICLE_ERPNR;
                            rowkit.price1 = mdpart.Properties.ARTICLE_SALESPRICE_1;
                            rowkit.price2 = mdpart.Properties.ARTICLE_SALESPRICE_2;
                        }
                        catch (NullReferenceException)
                        {
                            continue;
                        }

                    }
                    foreach (work KIT in MainList)
                    {
                        if (KIT.id == 0) continue;
                        foreach (part rowkit in KIT.parts)
                        {
                            foreach (part mainrowkit in MainList[0].parts)
                            {
                                if (mainrowkit.partnr == rowkit.partnr)
                                {
                                    rowkit.erpn = mainrowkit.erpn;
                                    rowkit.price1 = mainrowkit.price1;
                                    rowkit.price2 = mainrowkit.price2;
                                }

                            }
                        }
                    }
                    foreach (SpecPrice price in SpecALL)
                    {
                        foreach (work KIT in MainList)

                        {
                            if (price.id == KIT.id)
                            {
                                foreach (part rowkit in KIT.parts)
                                {
                                    price.Price1 += (rowkit.price1 * rowkit.pcount);
                                    price.Price2 += (rowkit.price2 * rowkit.pcount);

                                }
                                price.COSTPrice1 = price.Price1 * price.pcount;
                                price.COSTPrice2 = price.Price2 * price.pcount;
                            }
                        }
                    }
                    dataGridView4.Rows.Clear();
                    foreach (SpecPrice row in SpecALL)
                    {
                        dataGridView4.Rows.Add(
                            row.id,
                            row.NameKit,
                            row.pcount,
                            row.Price1,
                            row.Price2,
                            row.COSTPrice1,
                            row.COSTPrice2);
                    }
                }
                #endregion
                button3.Enabled = true;
                blockupdatelist = false;
                button1.Enabled = true;

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

            bw = new BackgroundWorker();
            if (bw.IsBusy) return;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork); 
            bw.RunWorkerAsync();
            button3.Enabled = false;
            
            //partup(1);


        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            //filtdevice2.Clear();
            dataGridView1.Rows.Clear();
            bool desloctrue = false;
            string buf = "";

            if (blockupdatelist == true) { return; }

            if (dataGridView2.CurrentRow.Index == 0)
                {
                filtdevice2 =  MainList[0].parts;
                //MessageBox.Show(dataGridView2.CurrentRow.Index.ToString()+"---"+ filtdevice2.Count.ToString());
                }
                else
                {
                    foreach(work f1 in MainList)
                {
                    if (dataGridView2[1, dataGridView2.CurrentRow.Index].Value.ToString() == f1.NameKit)
                    {
                        filtdevice2 = f1.parts;
                        //MessageBox.Show(dataGridView2.CurrentRow.Index.ToString() + "---" + filtdevice2.Count.ToString());
                        break;
                    }
                }
                    
                }

                for (int j = 0; j < filtdevice2.Count; j++)
                {
                   
                if (dataGridView2.CurrentRow.Index == 0)
                {
                    dataGridView1.Columns["Column4"].Visible = true;
                    /*dataGridView1.Columns["ERP"].Visible = true;
                    dataGridView1.Columns["Price1"].Visible = true;
                    dataGridView1.Columns["Price2"].Visible = true;*/
                    dataGridView1.Rows.Add(
                        j + 1,
                        filtdevice2[j].partnr,
                        filtdevice2[j].descrip,
                        filtdevice2[j].pcount,
                        filtdevice2[j].spare,
                        filtdevice2[j].erpn,
                        filtdevice2[j].price1,
                        filtdevice2[j].price2);
                    if (filtdevice2[j].spare < 0)
                    {
                        dataGridView1["Column4", dataGridView1.Rows.Count - 1].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        if (filtdevice2[j].spare == 0)
                        {
                            dataGridView1["Column4", dataGridView1.Rows.Count - 1].Style.BackColor = Color.Green;
                        }
                        else
                        {
                            dataGridView1["Column4", dataGridView1.Rows.Count - 1].Style.BackColor = Color.Yellow;
                        }
                            
                    }
                }
                
                else
                {
                    dataGridView1.Columns["Column4"].Visible = false;
                    /* dataGridView1.Columns["ERP"].Visible = false;
                     dataGridView1.Columns["Price1"].Visible = false;
                     dataGridView1.Columns["Price2"].Visible = false;*/
                    dataGridView1.Rows.Add(
                        j + 1,
                        filtdevice2[j].partnr,
                        filtdevice2[j].descrip,
                        filtdevice2[j].pcount,
                        "",
                        filtdevice2[j].erpn,
                        filtdevice2[j].price1,
                        filtdevice2[j].price2);
                }
                    
                }
            
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (blockupdatelist == true) { return; }
                dataGridView3.Rows.Clear();
                filtpart2.Clear();
                //filtpart2.Add(new part() { design = "Список использований", });
                bool searchpart = false;

                string bufname1 = "";
                string bufname2 = "";
                bufname1 = dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString();
                bufname1 = bufname1.Replace(" ", "");
                foreach (part p in filtpart)
                {
                    bufname2 = p.partnr;
                    bufname2 = bufname2.Replace(" ", "");
                    
                    if (bufname1 == bufname2)
                    {
                        searchpart = false;
                    foreach (part f2 in filtpart2)
                    {
                        //MessageBox.Show("--" + bufname1 + "--\n--" + bufname2 + "--");
                        if (p.design == f2.design)
                        {

                            searchpart = true;
                            f2.pcount += p.pcount;
                            break;
                        }
                    }
                            if (searchpart == false)
                            {
                                if (p.design=="")
                        {
                            filtpart2.Add(new part() { design = p.design, pcount = p.pcount, note = "Без структурных идентификаторов" });
                        }
                        else
                        {
                          filtpart2.Add(new part() { design = p.design, pcount = p.pcount,note=p.note });
                        }
                            
                            }
                        

                    }
                }

            
            for (int j=0;j<filtpart2.Count;j++)
            {
                dataGridView3.Rows.Add((j + 1), filtpart2[j].design, filtpart2[j].note, filtpart2[j].pcount);
            }
            int itog = 0;
            for (int i=0;i< dataGridView3.Rows.Count;i++)
            {  
                itog += Int32.Parse(dataGridView3[3, i].Value.ToString());
            }
            if (itog != 0)
            {
                dataGridView3.Rows.Add("", "Итого","", itog);
            }




        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
           
        }
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, ref int lpdwProcessId);
        private void button1_Click_3(object sender, EventArgs e)
        {

            openFileDialog1.Title = "Открыть файл шаблон";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string xlFileName = openFileDialog1.FileName;
            Excel.Application excl = new Excel.Application();
            Excel.Range Rng;
            Excel.Workbook xlWB;
            Excel.Worksheet xlSht;
            xlWB = excl.Workbooks.Open(xlFileName); //открываем наш файл
            long costprice1 = 0;
            long costprice2 = 0;
            const int Beginrow = 9; //Начальное число строк откуда начинать заполнение
            int startline = Beginrow;
            try
            {
                #region //Общий список устройств
                {
                    xlSht = xlWB.Worksheets["Общий Список Устройств"];
                    costprice1 = 0;
                    costprice2 = 0;
                    xlSht.Cells[1, "A"] = "Заказчик: " + CurProj.Properties.PROJ_ENDCUSTOMERNAME1;
                    xlSht.Cells[2, "A"] = "Название проекта: " + CurProj.ProjectName;
                    xlSht.Cells[3, "A"] = "Номер проекта: " + CurProj.Properties.PROJ_DRAWINGNUMBER;
                    xlSht.Cells[4, "A"] = "Название фирмы: " + CurProj.Properties.PROJ_COMPANYNAME;
                    xlSht.Cells[5, "A"] = "№ договора: ---";// + CurProj.Properties.PROJ_SUPPLEMENTARYFIELD;
                    xlSht.Cells[6, "A"] = "Автор: " + CurProj.Properties.PROJ_CREATOR;
                    xlSht.Cells[7, "A"] = "Последний обработчик: " + CurProj.Properties.PROJ_LASTMODIFICATOR;
                    startline = Beginrow;
                    //Rng=xlSht.get_Range(xlSht.Cells[i + startline, 0], xlSht.Cells[i + startline, 7]);
                    Rng = (Excel.Range)xlSht.get_Range("A" + (startline).ToString(), "L" + (startline).ToString()).Cells;
                    Rng.Merge(Type.Missing);
                    xlSht.Cells[startline, "A"] = MainList[0].NameKit;
                    xlSht.Cells[startline, "A"].Font.Size = 12;
                    xlSht.Cells[startline, "A"].Font.Bold = true;
                    xlSht.Cells[startline, "A"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[startline, "A"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    startline++;
                    for (int j = 0; j < MainList[0].parts.Count; j++)
                    {
                        xlSht.Cells[startline, "A"] = (j + 1).ToString();
                        xlSht.Cells[startline, "B"] = MainList[0].parts[j].partnr;
                        xlSht.Cells[startline, "C"] = MainList[0].parts[j].pcount;
                        xlSht.Cells[startline, "D"] = MainList[0].parts[j].spare;

                        if (MainList[0].parts[j].spare < 0)
                        {
                            xlSht.Cells[startline, "D"].Interior.Color = ColorTranslator.ToOle(Color.Red);
                        }
                        else
                        {
                            if (MainList[0].parts[j].spare == 0)
                            {
                                xlSht.Cells[startline, "D"].Interior.Color = ColorTranslator.ToOle(Color.Green);
                            }
                            else
                            {
                                xlSht.Cells[startline, "D"].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                            }

                        }
                        xlSht.Cells[startline, "E"] = MainList[0].parts[j].pcount - MainList[0].parts[j].spare;
                        xlSht.Cells[startline, "F"] = "шт";
                        xlSht.Cells[startline, "G"] = MainList[0].parts[j].descrip;
                        xlSht.Cells[startline, "H"] = MainList[0].parts[j].erpn;
                        xlSht.Cells[startline, "I"] = MainList[0].parts[j].price1;
                        xlSht.Cells[startline, "J"] = MainList[0].parts[j].price2;
                        xlSht.Cells[startline, "K"] = MainList[0].parts[j].price1 * MainList[0].parts[j].pcount;
                        xlSht.Cells[startline, "L"] = MainList[0].parts[j].price2 * MainList[0].parts[j].pcount;
                        costprice1 += (MainList[0].parts[j].price1 * MainList[0].parts[j].pcount);
                        costprice2 += (MainList[0].parts[j].price2 * MainList[0].parts[j].pcount);
                        startline++;

                    }
                    Rng = xlSht.get_Range("A" + (startline).ToString(), "I" + (startline).ToString()).Cells;
                    Rng.Merge(Type.Missing);
                    xlSht.Cells[startline, "J"] = "Итого:";
                    xlSht.Cells[startline, "J"].Font.Bold = true;
                    xlSht.Cells[startline, "J"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[startline, "J"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    xlSht.Cells[startline, "K"] = costprice1;
                    //xlSht.Cells[i + startline, "H"].NumberFormat = "#,###.00 ₸";
                    xlSht.Cells[startline, "K"].Font.Bold = true;
                    xlSht.Cells[startline, "K"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[startline, "K"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    xlSht.Cells[startline, "L"] = costprice2;
                    //xlSht.Cells[i + startline, "I"].NumberFormat = "#,###.00 €";
                    xlSht.Cells[startline, "L"].Font.Bold = true;
                    xlSht.Cells[startline, "L"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[startline, "L"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Rng = xlSht.get_Range("A" + 4, "L" + (startline).ToString()).Cells;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                    //startline++;

                }
                #endregion

                #region //Список устройств

                {
                    xlSht = xlWB.Worksheets["Список Устройств"];
                    startline = Beginrow;

                    for (int i = 1; i < MainList.Count; i++)
                    {
                        costprice1 = 0;
                        costprice2 = 0;
                        //Rng=xlSht.get_Range(xlSht.Cells[i + startline, 0], xlSht.Cells[i + startline, 7]);
                        Rng = (Excel.Range)xlSht.get_Range("A" + (startline).ToString(), "J" + (startline).ToString()).Cells;
                        Rng.Merge(Type.Missing);
                        xlSht.Cells[startline, "A"] = MainList[i].NameKit;
                        xlSht.Cells[startline, "A"].Font.Size = 12;
                        xlSht.Cells[startline, "A"].Font.Bold = true;
                        xlSht.Cells[startline, "A"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        xlSht.Cells[startline, "A"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        startline++;
                        for (int j = 0; j < MainList[i].parts.Count; j++)
                        {

                            xlSht.Cells[startline, "A"] = (j + 1).ToString();
                            xlSht.Cells[startline, "B"] = MainList[i].parts[j].partnr;
                            xlSht.Cells[startline, "C"] = MainList[i].parts[j].pcount;
                            xlSht.Cells[startline, "D"] = "шт";
                            xlSht.Cells[startline, "E"] = MainList[i].parts[j].descrip;
                            xlSht.Cells[startline, "F"] = MainList[i].parts[j].erpn;
                            xlSht.Cells[startline, "G"] = MainList[i].parts[j].price1;
                            xlSht.Cells[startline, "H"] = MainList[i].parts[j].price2;
                            xlSht.Cells[startline, "I"] = MainList[i].parts[j].price1 * MainList[i].parts[j].pcount;
                            xlSht.Cells[startline, "J"] = MainList[i].parts[j].price2 * MainList[i].parts[j].pcount;
                            costprice1 += (MainList[i].parts[j].price1 * MainList[i].parts[j].pcount);
                            costprice2 += (MainList[i].parts[j].price2 * MainList[i].parts[j].pcount);
                            startline++;
                        }
                        Rng = xlSht.get_Range("A" + (startline).ToString(), "G" + (startline).ToString()).Cells;
                        Rng.Merge(Type.Missing);
                        xlSht.Cells[startline, "H"] = "Итого:";
                        xlSht.Cells[startline, "H"].Font.Bold = true;
                        xlSht.Cells[startline, "H"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        xlSht.Cells[startline, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        xlSht.Cells[startline, "I"] = costprice1;
                        xlSht.Cells[startline, "I"].Font.Bold = true;
                        xlSht.Cells[startline, "I"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        xlSht.Cells[startline, "I"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        xlSht.Cells[startline, "J"] = costprice2;
                        //xlSht.Cells[i + startline, "I"].NumberFormat = "#,###.00 €";
                        xlSht.Cells[startline, "J"].Font.Bold = true;
                        xlSht.Cells[startline, "J"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        xlSht.Cells[startline, "J"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        Rng = xlSht.get_Range("A" + 4, "J" + (startline).ToString()).Cells;
                        Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                        Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                        Rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                        Rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                        Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                        startline++;


                    }

                }
                #endregion

                #region //КП
                {
                    xlSht = xlWB.Worksheets["КП"];
                    startline = Beginrow;
                    costprice1 = 0;
                    costprice2 = 0;
                    for (int i = 0; i < SpecALL.Count; i++)
                    {
                        if (SpecALL[i].id == 0) continue;
                        xlSht.Cells[i + startline, "A"] = SpecALL[i].id;
                        xlSht.Cells[i + startline, "B"] = SpecALL[i].NameKit;
                        xlSht.Cells[i + startline, "C"] = SpecALL[i].pcount;
                        xlSht.Cells[i + startline, "D"] = "компл";
                        xlSht.Cells[i + startline, "E"] = SpecALL[i].Price1;
                        xlSht.Cells[i + startline, "F"] = SpecALL[i].Price2;
                        xlSht.Cells[i + startline, "G"] = SpecALL[i].COSTPrice1;
                        xlSht.Cells[i + startline, "H"] = SpecALL[i].COSTPrice2;
                        costprice1 += SpecALL[i].COSTPrice1;
                        costprice2 += SpecALL[i].COSTPrice2;

                    }
                    xlSht.Cells[SpecALL.Count + startline, "F"] = "Итого:";
                    xlSht.Cells[SpecALL.Count + startline, "F"].Font.Bold = true;
                    xlSht.Cells[SpecALL.Count + startline, "F"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[SpecALL.Count + startline, "F"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    xlSht.Cells[SpecALL.Count + startline, "G"] = costprice1;
                    xlSht.Cells[SpecALL.Count + startline, "G"].Font.Bold = true;
                    xlSht.Cells[SpecALL.Count + startline, "G"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[SpecALL.Count + startline, "G"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    xlSht.Cells[SpecALL.Count + startline, "H"] = costprice2;
                    xlSht.Cells[SpecALL.Count + startline, "H"].Font.Bold = true;
                    xlSht.Cells[SpecALL.Count + startline, "H"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlSht.Cells[SpecALL.Count + startline, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Rng = xlSht.get_Range("A" + 4, "H" + (SpecALL.Count + startline).ToString()).Cells;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                    Rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                }
                #endregion

            }
            catch
            {
                MessageBox.Show("Ошибка при работе с файлом, возможно шаблон файла не актуален.");
                xlWB.Close(false); //сохраняем и закрываем файл
                excl.Quit();
                return;
            }
            
            saveFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
            saveFileDialog1.Title = "Сохранить Данные в файл Excel";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            excl.Application.ActiveWorkbook.SaveAs(saveFileDialog1.FileName, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWB.Close(false); //сохраняем и закрываем файл
            excl.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSht);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWB);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excl);
            int tExcelPID = 0;
            int tHwnd = 0;
            tHwnd = excl.Hwnd; //Получим HWND окна
            System.Diagnostics.Process tExcelProcess;
            GetWindowThreadProcessId((IntPtr)tHwnd, ref tExcelPID); //По HWND получим PID
            tExcelProcess = System.Diagnostics.Process.GetProcessById(tExcelPID); //Подключимся к процессу                                               ////Убийство процесса Excel
            tExcelProcess.Kill();
            tExcelProcess = null;
            xlSht = null;
            xlWB = null;
            excl = null;
            System.GC.Collect();
        }
        /*public void SerializeAndSave(string path, List<work> data)
        {
            var serializer = new XmlSerializer(typeof(List<work>));
            using (var writer = new StreamWriter(path))
            {
                serializer.Serialize(writer, data);
            }
        }
        public void SerializeAndSave(string path, List<SpecPrice> data)
        {
            var serializer = new XmlSerializer(typeof(List<SpecPrice>));
            using (var writer = new StreamWriter(path))
            {
                serializer.Serialize(writer, data);
            }
        }*/
    }
}
