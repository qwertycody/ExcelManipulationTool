using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelManipulationTool_Launcher
{
    class Program
    {
        static void Main(string[] args)
        {
            String directoryToOutputTo = Path.GetTempPath() + Path.GetRandomFileName();

            if (!Directory.Exists(directoryToOutputTo))
            {
                Directory.CreateDirectory(directoryToOutputTo);
            }

            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ClosedXML.dll"), Properties.Resources.ClosedXML_DLL);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ClosedXML.xml"), Properties.Resources.ClosedXML_XML);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "DocumentFormat.OpenXml.dll"), Properties.Resources.DocumentFormat_OpenXml_DLL);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "DocumentFormat.OpenXml.xml"), Properties.Resources.DocumentFormat_OpenXml_XML);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ExcelNumberFormat.dll"), Properties.Resources.ExcelNumberFormat_DLL);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ExcelNumberFormat.xml"), Properties.Resources.ExcelNumberFormat_XML);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ExcelManipulationTool.exe"), Properties.Resources.ExcelManipulationTool_EXE);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ExcelManipulationTool.pdb"), Properties.Resources.ExcelManipulationTool_PDB);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "ExcelManipulationTool.exe.config"), Properties.Resources.ExcelManipulationTool_EXE_CONFIG);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "FastMember.Signed.dll"), Properties.Resources.FastMember_Signed_DLL);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "System.IO.FileSystem.Primitives.dll"), Properties.Resources.System_IO_FileSystem_Primitives_DLL);
            File.WriteAllBytes(getWriteString(directoryToOutputTo, "System.IO.Packaging.dll"), Properties.Resources.System_IO_Packaging_DLL);

            //File.WriteAllBytes(getWriteString(directoryToOutputTo, "GemBox.Spreadsheet.dll"), Properties.Resources.GemBox_Spreadsheet_DLL);
            //File.WriteAllBytes(getWriteString(directoryToOutputTo, "GemBox.Spreadsheet.xml"), Properties.Resources.GemBox_Spreadsheet_XML);


            var process = Process.Start(getWriteString(directoryToOutputTo, "ExcelManipulationTool.exe"));
            process.WaitForExit();

            Directory.Delete(directoryToOutputTo, true);
        }

        private static String getWriteString(String directory, String filename)
        {
            return directory + @"\" + filename;
        }
    }
}
