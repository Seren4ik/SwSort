using CodeStack.SwEx.AddIn;
using CodeStack.SwEx.AddIn.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.IO;
using ExcelDataReader;
using System.Data;

namespace SwSort
{
    [ComVisible(true), Guid("7E29F976-F537-4530-BE05-39FD0CA92213")]
    [AutoRegister("SwSortament", "Search Sortament")]


    public class SwSort : SwAddInEx
    {
        private enum Commands_e
        {
            SearchSortament
        }

        public override bool OnConnect()
        {
            AddCommandGroup<Commands_e>(OnButtonClick);
            return true;
        }

        private void OnButtonClick(Commands_e cmd)
        {
            switch (cmd)
            {
                case Commands_e.SearchSortament:
                    SearchSortament1();
                    break;
            }
        }

        private void SearchSortament1()
        {
            
            string path1 = @"d:\Для Конструкторов\Сортамент1.xlsx";
            IModelDoc2 swModel;
            IModelDoc2 swModel2;
            Component2 swComponent;
            Object[] Components;
            AssemblyDoc swAssembly;

            swModel = (IModelDoc2)App.ActiveDoc;

            if (swModel is AssemblyDoc)
            {
                swAssembly = (AssemblyDoc)swModel;
                Components = (Object[])swAssembly.GetComponents(false);
                foreach (Object SingleComponent in Components)
                {
                    swComponent = (Component2)SingleComponent;             
                    string pathcomp = swComponent.GetPathName();
                    

                    if (pathcomp.Contains("Саморез") || pathcomp.Contains("Шуруп") || pathcomp.Contains("Болт") || pathcomp.Contains("Винт") || pathcomp.Contains("Гайка") || pathcomp.Contains("Шайба"))
                    {
                        App.SendMsgToUser2("Метизы", 0, 0);
                    }
                    else
                    {                        
                        //swModel2 = (IModelDoc2)App.IActivateDoc(pathcomp);
                        swModel2 = (IModelDoc2)swComponent.GetModelDoc2();
                       // GetInfo(swModel2,path1);
                        GetInfo(swModel2, path1);
                    }
                }
                
            }
            else
            {
                GetInfo(swModel, path1);
            }
        }


        private void GetInfo(IModelDoc2 swModel, string path1)
        {
            //string[] listConf = (string[])swModel.GetConfigurationNames();
            string nameConf = swModel.IGetActiveConfiguration().Name.ToString();

            //foreach (string nameConf in listConf)
            //{


            string type1 = swModel.CustomInfo["Тип"];
            string tolsh1 = swModel.GetCustomInfoValue(nameConf, "Толщина");
            string widht = swModel.GetCustomInfoValue(nameConf, "Ширина");
            string tolshtr = swModel.GetCustomInfoValue(nameConf, "Толщина трубы");
            string volume = swModel.GetCustomInfoValue(nameConf, "Объём");
            string surface_area = swModel.GetCustomInfoValue(nameConf, "Площадь поверхности");

            #region Добавление объёма  площади
            if (volume == "")
            {
                swModel.AddCustomInfo3(nameConf,"Объём",1, "\"SW-Volume@@По умолчанию@Деталь1.SLDPRT\"");
            }

            if (surface_area == "")
            {
                swModel.AddCustomInfo3(nameConf, "Площадь поверхности", 1, "\"SW-Площадь поверхности@@По умолчанию@Деталь1.SLDPRT\"");
            }
            #endregion
            #region Лист
            if (type1 == "Лист" || type1 == "Лист ОЦ")
                {
                    if (tolsh1 == "1" || tolsh1 == "2" || tolsh1 == "3" || tolsh1 == "4" || tolsh1 == "5" || tolsh1 == "6" || tolsh1 == "8" || tolsh1 == "10" || tolsh1 == "12")
                    {
                        tolsh1 = tolsh1 + ",0";
                    }

                    string typesort = type1 + " Б-ПН-" + tolsh1;

                    OpenExcelFile(path1, typesort, nameConf, swModel);
                }
            else if (type1.Contains("Лист ПВЛ"))
            {
                if (tolsh1 == "1" || tolsh1 == "2" || tolsh1 == "3" || tolsh1 == "4" || tolsh1 == "5" || tolsh1 == "6" || tolsh1 == "8" || tolsh1 == "10" || tolsh1 == "12")
                {
                    tolsh1 = tolsh1 + ",0";
                }

                string typesort = type1 + " TR16-ОЦ-" + tolsh1;

                OpenExcelFile(path1, typesort, nameConf, swModel);
            }    
                #endregion
                #region Труба Уголок
                else if (type1 == "Труба" || type1 == "Труба проф." || type1 == "Уголок")
                {
                    if (type1 == "Уголок")
                    {
                        string typesort1 = swModel.CustomInfo["description"].Replace(" ", "");
                        if (typesort1 != "")
                        {
                        App.SendMsgToUser2(typesort1, 0, 0);
                        string typesort3 = type1 + " " + "В-" + typesort1.Replace(".0", "").Replace("x", "х");
                            OpenExcelFile(path1, typesort3, nameConf, swModel);
                        }
                        else
                        {
                        App.SendMsgToUser2("description не обнаружен", 0, 0);
                            string typesort = type1 + " " + "В-" + tolsh1 + "х" + widht + "х" + tolshtr;
                            OpenExcelFile(path1, typesort, nameConf, swModel);
                        }

                    }
                    else
                    {
                        string typesort1 = swModel.CustomInfo["description"].Replace(" ", "");
                        if (typesort1 != "")
                        {
                        App.SendMsgToUser2(typesort1, 0, 0);
                        string typesort3 = type1 + " " + typesort1.Replace(".", ",").Replace("x", "х");
                            OpenExcelFile(path1, typesort3, nameConf, swModel);
                        }
                        else
                        {
                        App.SendMsgToUser2("description не обнаружен", 0, 0);
                        string typesort = type1 + " " + tolsh1 + "х" + widht + "х" + tolshtr;
                            OpenExcelFile(path1, typesort, nameConf, swModel);
                        }
                    }
                }

                else if (type1 == "Труба круг." || type1 == "Труба ВГП.")
                {
                    string typesort = type1 + " " + tolsh1 + "х" + tolshtr;

                    OpenExcelFile(path1, typesort, nameConf, swModel);

                }
                #endregion
                #region Полоса Квадрат Брусок Доска
                else if (type1 == "Полоса" || type1 == "Квадрат" || type1 == "Брусок" || type1 == "Доска" || type1 == "Брус")
                {
                    if (type1 == "Квадрат")
                    {
                        string typesort = type1 + " " + "В1-" + widht + "х" + tolsh1;

                        OpenExcelFile(path1, typesort, nameConf, swModel);
                    }
                    else if (type1 == "Полоса")
                    {
                        string typesort = type1 + " " + widht + "х" + tolsh1;

                        OpenExcelFile(path1, typesort, nameConf, swModel);
                    }
                    else
                    {
                        string typesort = type1 + " " + tolsh1 + "х" + widht;

                        OpenExcelFile(path1, typesort, nameConf, swModel);
                    }

                }
                #endregion
                #region Круг Арматура
                else if (type1 == "Круг" || type1 == "Арматура период." || type1 == "Арматура глад.")
                {
                    string typesort = type1 + " " + tolsh1;

                    OpenExcelFile(path1, typesort, nameConf, swModel);
                }
                #endregion
                else
                {
                    App.SendMsgToUser2("Тип не определен", 0, 0);
                    _ = App == null;

                }
           // }
        }

        //Чтение Excel файла 
        private void OpenExcelFile(string path, string typesort, string nameConf, IModelDoc2 SW)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            DataSet result = reader.AsDataSet();

            DataTable dt = result.Tables[0];

            stream.Close();
            string[] listProp = new string[2];


            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    string Text = Convert.ToString(dt.Rows[i][j]);

                    //Проверка на наличие в спецификации Excel 
                    if (Text.Contains(typesort))
                    {

                        listProp[0] = Convert.ToString(dt.Rows[i][j]);
                        listProp[1] = Convert.ToString(dt.Rows[i][j + 1]);
                    }

                }
            }


            #region Создание свойств 
            if (listProp[0] != null)
            {               
                string[] prop1 = listProp[0].Split(' ');
               
                string namedoc = SW.GetPathName();
                SW.DeleteCustomInfo2(nameConf,"Сортамент");
                SW.DeleteCustomInfo2(nameConf,"Сорт");
                SW.DeleteCustomInfo2(nameConf,"Материал");
                SW.DeleteCustomInfo("Сортамент");
                SW.AddCustomInfo("Сортамент", "Текст", listProp[0]);
                SW.AddCustomInfo2("Сортамент", 1, listProp[0]);
                if (prop1.Length == 4)
                {
                    SW.DeleteCustomInfo("Сорт");
                    SW.AddCustomInfo("Сорт", "Текст", prop1[1] + " " + prop1[2] + " " + prop1[3]);
                    SW.AddCustomInfo2("Сорт", 1, prop1[1] + " " + prop1[2] + " " + prop1[3]);
                }
                else if (prop1.Length == 5)
                {
                    SW.DeleteCustomInfo("Сорт");
                    SW.AddCustomInfo("Сорт", "Текст",prop1[2] + " " + prop1[3]+ " " + prop1[4]);
                    SW.AddCustomInfo2("Сорт", 1,prop1[2] + " " + prop1[3] + " " + prop1[4]);
                }
                else
                {
                    SW.DeleteCustomInfo("Сорт");
                    SW.AddCustomInfo("Сорт", "Текст", prop1[1]);
                    SW.AddCustomInfo2("Сорт", 1, prop1[1]);

                }
                
                SW.DeleteCustomInfo("Материал");
                SW.AddCustomInfo("Материал", "Текст", listProp[1]);
                SW.AddCustomInfo2("Материал", 1, listProp[1]);
                SW.EditRebuild3();
                SW.Save();
                App.SendMsgToUser2("Операция выполнена", 2, 0);
                GC.Collect();

            }

            else if (listProp[0] == null)
            {

                SW.DeleteCustomInfo("Сортамент");
                SW.AddCustomInfo("Сортамент", "Текст", "Сортамент не найден");
                SW.AddCustomInfo2("Сортамент", 1, "Сортамент не найден");

                SW.EditRebuild3();
                App.SendMsgToUser2("Сортамент не найден", 0, 0);
                GC.Collect();

            }

        }
        #endregion


    }

}