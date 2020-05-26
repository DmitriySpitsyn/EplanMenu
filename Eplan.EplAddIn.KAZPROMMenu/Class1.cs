using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;
using Eplan.EplSDK.WPF;
using System.Reflection;
using Eplan.EplSDK.WPF.Interfaces;
using Eplan.EplSDK.WPF.Interfaces.DialogServices;
using Eplan.EplApi.System;
using System;

namespace Class1
{ 
    public class AddInModule : IEplAddIn
    {
        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        public bool OnUnregister()
        {
            return true;
        }

        public bool OnInit()
        {
            return true;
        }

        public bool OnInitGui()
        {
            Menu OurMenu = new Menu();
            OurMenu.AddMainMenu("KAZPROM ENGINEERING", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuUtilities, "Обновить цены", "UpdatePrice","Обновить цены с 1С", 1);
            OurMenu.AddPopupMenuItem("Работа с устройствами", "Проверить расположение в 3D", "Test3dver2", "Проверка расположения устройства в виде 3D", OurMenu.GetCustomMenuId("Обновить цены", "UpdatePrice"), 1, false, false);
             OurMenu.AddMenuItem("О проекте", "About", "Информация о проекте", OurMenu.GetCustomMenuId("Обновить цены", "UpdatePrice"), 2, true, true);
            OurMenu.AddMenuItem("Работа с спецификацией", "PLC", "Работа с спецификацией", OurMenu.GetCustomMenuId("Проверить расположение в 3D", "Test3dver2"), 2, false, false);
           // OurMenu.AddMenuItem("Расположение в 3D 2", "Test3dver2", "Проверить расположение в 3D2", OurMenu.GetCustomMenuId("Работа с спецификацией", "PLC"), 2, false, false);
            return true;
        }

        public bool OnExit()
        {
            return true;
        }
       
        }
    }
