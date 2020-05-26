using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.HEServices;
using System.Windows.Forms;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    class Placement3D
    {
        public class Action_Test : IEplAction
        {
            public bool OnRegister(ref string Name, ref int Ordinal)
            {
                Name = "Placement3D";
                Ordinal = 20;
                return true;
            }
            public bool Execute(ActionCallingContext oActionCallingContext)
            {
                
                Form1 newForm = new  Form1();
                newForm.Show();
                return true;
                
                


            }

            public void GetActionProperties(ref ActionProperties actionProperties)
            {

            }


        }

    }
}

