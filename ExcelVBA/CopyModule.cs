using System;
using VBIDE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using VBA = Microsoft.Vbe.Interop;


namespace ExcelVBA
{
    class CopyModule
    {
        public static Hashtable GetMacros(Excel._Workbook wDestination)
        {
            Hashtable pak = new Hashtable();
            VBA.VBProject prj;
            VBA.CodeModule code;
            string composedFile;
            prj = wDestination.VBProject;

            foreach (VBA.VBComponent comp in prj.VBComponents)
            {

                if (comp.Type == VBA.vbext_ComponentType.vbext_ct_StdModule)
                {
                    code = comp.CodeModule;


                    composedFile = comp.Name + Environment.NewLine;


                    for (int i = 0; i < code.CountOfLines; i++)
                    {
                        composedFile +=
                            code.get_Lines(i + 1, 1) + Environment.NewLine;
                    }
                    pak.Add(comp.Name, composedFile);
                }
            }

            return pak;
        }


        public void CopyMacro(string path1, string path2)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
          
            Excel._Workbook wSource = app.Workbooks.Open(path1), wDestination = app.Workbooks.Open(path2);
            
            bool IsSourceProtected = Convert.ToBoolean(wSource.VBProject.Protection);
            bool IsDestinationProtected = Convert.ToBoolean(wDestination.VBProject.Protection);
            if (IsSourceProtected)
            {
                if (IsDestinationProtected)
                {
                    KeySendPassword pasike = new KeySendPassword();
                    pasike.Klucze();
                }
                else
                {
                    KeySendPassword pasik = new KeySendPassword();
                    pasik.KluczS(ref app);
                }
            }
            else
           if (IsDestinationProtected)
            {
                KeySendPassword pasiks = new KeySendPassword();
                pasiks.KluczD();
            }
            Boolean found;
            found = false;
            VBIDE.VBComponent dest;

            foreach (VBIDE.VBComponent source in wSource.VBProject.VBComponents)
            {
                //Sprawdzamy czy nasz source ma jakis kod jezeli nie: koniec.
                if (source.CodeModule.CountOfLines > 0)
                {
                    //Sprawdzamy czy istnieje jakies makro w naszym destini pliku jezeli tak sprawdzamy i porwonujemy jego nazwe jezeli nie : dalej.
                    Hashtable pak = new Hashtable();
                    pak = GetMacros(wDestination);
                    if (pak.Count > 0)
                    {
                        //Sprawdzamy czy dany modul istnieje
                        //I czy jego nazwa jest taka sama jak z sourca jezeli tak to jest kasowana jezeli nie to zostaje.

                        foreach (VBIDE.VBComponent destNew in wDestination.VBProject.VBComponents)
                        {
                            if (destNew.Name == source.CodeModule.Name)
                            {
                                wDestination.VBProject.VBComponents.Remove(destNew);
                                found = false;

                                //Usuwamy ten sam module
                            }

                        }
                    }
                    if (found == false)
                    {

                        dest = wDestination.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
                        dest.CodeModule.AddFromString(source.CodeModule.get_Lines(1, source.CodeModule.CountOfLines));

                        dest.Name = source.Name;
                        wDestination.Save();
                       
                        Marshal.FinalReleaseComObject(dest);
                        Marshal.FinalReleaseComObject(source);
                        dest = null;
                    }

                }
            }
           
            wSource.Close();
            wDestination.Close();
            Marshal.FinalReleaseComObject(wSource);
            Marshal.FinalReleaseComObject(wDestination);
            app.Quit();
        }
    }
}