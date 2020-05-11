using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.HEServices;
using System.Windows.Forms;
using Eplan.EplApi.DataModel.E3D;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    class About
    {
        public class Action_Test : IEplAction
        {
            public bool OnRegister(ref string Name, ref int Ordinal)
            {
                Name = "About";
                Ordinal = 20;
                return true;
            }

            public bool Execute(ActionCallingContext oActionCallingContext)
            {
                SelectionSet Set = new SelectionSet();
                Project CurrentProject = Set.GetCurrentProject(true);
                
                string ProjectName = CurrentProject.ProjectName;
                string ProjectCompanyName = CurrentProject.Properties.PROJ_COMPANYNAME;
                DateTime ProjectCreationDate = CurrentProject.Properties.PROJ_CREATIONDATE;
                MessageBox.Show("Название проекта: " + ProjectName + "\n" + "Название фирмы: " + ProjectCompanyName +
                                "\n" + "Дата создания проекта: " + ProjectCreationDate.ToShortDateString()+"\n"+"Автор API: Спицын Дмитрий.","О проекте");
               
                return true;
            }

            public void GetActionProperties(ref ActionProperties actionProperties)
            {
            }
        }
    }
}
