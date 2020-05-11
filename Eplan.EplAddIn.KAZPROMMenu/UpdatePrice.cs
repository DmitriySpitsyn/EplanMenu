using System;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;


namespace Eplan.EplAddIn.KAZPROMMenu
{
    public class Action_Test : IEplAction
    {
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "UpdatePrice";
            Ordinal = 20;
            return true;
        }

        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            UpdatePriceForm newform = new UpdatePriceForm();
            newform.Show();
            
            return true;
        }

        public void GetActionProperties(ref ActionProperties actionProperties)
        {
        }
    }


}

