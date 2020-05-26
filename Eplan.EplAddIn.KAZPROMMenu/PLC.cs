using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eplan.EplApi.ApplicationFramework;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    class PLC
    {
        public class Action_Test : IEplAction
        {
            public bool OnRegister(ref string Name, ref int Ordinal)
            {
                Name = "PLC";
                Ordinal = 20;
                return true;
            }
            public bool Execute(ActionCallingContext oActionCallingContext)
            {
                Form2 newForm = new Form2();
                newForm.Show();
                return true;




            }

            public void GetActionProperties(ref ActionProperties actionProperties)
            {

            }

        }
    }
}
    
