using System;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace RenameSpreadsheets
{
    class XL
    {
        public static void RenameSpreadsheets()
        {
            Console.Write("Renomear Planilhas? (s/n): ");
            char continuar = Console.ReadKey().KeyChar;
            Console.WriteLine();

            if (char.ToLower(continuar) != 's')
            {
                Quit();
            }

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var inDefaultFolder = desktop + @"\RenomearPlanilhas";
            var outDefaultFolder = desktop + @"\PlanilhasRenomeadas";

            if (!ContainsFiles(inDefaultFolder))
            {
                Console.WriteLine("Diretorio Vazio");
                Quit();
            };

            var exit = "";

            var outMensal = outDefaultFolder + @"\Mensal";
            var outBimestral = outDefaultFolder + @"\Bimestral";
            var outICEP = outDefaultFolder + @"\ICEP";
            var outSAEB = outDefaultFolder + @"\SAEB";

            var xl = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            foreach (var fileStr in GetFilesStr(inDefaultFolder))
            {
                var fileinfo = new FileInfo(fileStr);
                var wb = xl.Workbooks.Open(fileinfo.FullName);

                var outFileStr = "";
                var school = "";
                
                try
                {
                    // Get the control spreadsheet
                    var ws = (Excel.Worksheet)wb.Sheets[3];
                    ws.Unprotect("3dUc@c@0");

                    int type = (int) ((Excel.Range)ws.Cells[4, 1]).Value;

                    // Verification, because ICEP have different school name cell position
                    if (type == (int) Type.ICEP)
                    {
                        school = ((Excel.Range)ws.Cells[17, 2]).Value.ToString();
                    }
                    else
                    {
                        school = ((Excel.Range)ws.Cells[18, 2]).Value.ToString();
                    }

                    outFileStr = school + fileinfo.Extension;

                    var isEmpty = school == "Digite o Código da Escola no campo acima!";

                    if (isEmpty)
                    {
                        Console.WriteLine("[EscolaId não preenchido] " + school + " - Nome Original: " + fileinfo.Name);
                        wb.Close();
                        continue;
                    }

                    if (type == (int) Type.Mensal)
                    {
                        if (!Directory.Exists(outMensal))
                        {
                            Directory.CreateDirectory(outMensal);
                        }

                        outFileStr = school + " - [Mensal]" + fileinfo.Extension;

                        exit = outMensal;
                    }
                    
                    if (type == (int) Type.Bimestral)
                    {
                        if (!Directory.Exists(outBimestral))
                        {
                            Directory.CreateDirectory(outBimestral);
                        }

                        outFileStr = school + fileinfo.Extension;

                        exit = outBimestral;
                    }
                    
                    if (type == (int) Type.ICEP)
                    {
                        if (!Directory.Exists(outICEP))
                        {
                            Directory.CreateDirectory(outICEP);
                        }

                        outFileStr = school + " - [ICEP]" + fileinfo.Extension;

                        exit = outICEP;
                    }
                    
                    if (type == (int) Type.SAEB)
                    {
                        if (!Directory.Exists(outSAEB))
                        {
                            Directory.CreateDirectory(outSAEB);
                        }

                        outFileStr = school + " - [SAEB]" + fileinfo.Extension;

                        exit = outSAEB;
                    }
                }
                catch
                {
                    Console.WriteLine("[Erro ao abrir] " + school + " - Nome Original: " + fileinfo.Name);
                }

                wb.Close();

                try
                {
                    File.Move(fileinfo.FullName, exit + @"\" + outFileStr);
                    Console.WriteLine(outFileStr);
                }
                catch (Exception e)
                {
                    Console.WriteLine("[Erro] " + school + " - Nome Original: " + fileinfo.Name + " | " + e.Message);
                    Console.WriteLine("Pressione qualquer tecla para continuar!");
                    Console.ReadKey();
                }
            }
            xl.Quit();
        }

        private static bool ContainsFiles(string inDefaultFolder)
        {
            return Directory.GetFiles(inDefaultFolder, "*.xls*").Where(x => !x.Contains("$")).Any();
        }

        private static string[] GetFilesStr(string inDefaultFolder)
        {
            return Directory.GetFiles(inDefaultFolder, "*.xls*").Where(x => !x.Contains("$")).ToArray();
        }

        public static void Quit()
        {
            Console.WriteLine("Pressione qualquer tecla para sair!");
            Console.ReadKey();
            Environment.Exit(0);
        }

        private enum Type
        {
            Mensal = 0,
            Bimestral = 1,
            ICEP = 2,
            SAEB = 3
        }
    }
}
