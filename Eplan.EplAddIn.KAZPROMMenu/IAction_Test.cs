using Eplan.EplApi.ApplicationFramework;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    public interface IAction_Test
    {
        bool Execute(ActionCallingContext oActionCallingContext);
        void GetActionProperties(ref ActionProperties actionProperties);
        bool OnRegister(ref string Name, ref int Ordinal);
    }
}