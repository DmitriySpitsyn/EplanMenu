using Eplan.EplApi.ApplicationFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    class Test3dver2
    {
        public class Action_Test : IEplAction
        {
            public bool OnRegister(ref string Name, ref int Ordinal)
            {
                Name = "Test3dver2";
                Ordinal = 20;
                return true;
            }
            public bool Execute(ActionCallingContext oActionCallingContext)
            {
                Form1ver2 newForm = new Form1ver2();
                newForm.Show();
                return true;

            }
            public void GetActionProperties(ref ActionProperties actionProperties)
            {

            }

        }
    }
}
