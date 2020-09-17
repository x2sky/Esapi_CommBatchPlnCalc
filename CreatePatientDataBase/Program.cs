//////////////////////////////////////////////////////////////////////
///This program copy plan from multiple reference patients and paste to an archived database patient for beam commissioning purpose
///version 0.0
///Becket Hui 2020/9
//////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using VMS.TPS.Common.Model.API;
using System.Windows;
using System.Windows.Forms;
using Application = VMS.TPS.Common.Model.API.Application;
using System.IO;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("0.0.0.0")]
[assembly: AssemblyFileVersion("0.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace CreatePatientDataBase
{
    class Program
    {
        public static string msg;
        [STAThread]
        static void Main(string[] args)
        {
            //Open file dialog box to select list of IMRT QA plans//
            var infile = new OpenFileDialog
            {
                Multiselect = false,
                Title = "Choose IMRT QA plans list",
                Filter = "Comma-separated values file|*.csv"
            };
            //Open file dialog box to select archived information//
            var outfile = new SaveFileDialog
            {
                Title = "Choose the file to save the database",
                Filter = "Comma-separated values file|*.csv",
                OverwritePrompt = false
            };
            if (infile.ShowDialog() == DialogResult.OK && outfile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    //Start application//
                    Console.WriteLine("Start to create patient database...\n");
                    using (Application app = Application.CreateApplication())
                    {
                        Execute(app, infile.FileName, outfile.FileName);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.ToString());
                }
                Console.WriteLine("Database is created. Press any key to exit.");
                Console.ReadKey();
            }
        }
        static void Execute(Application app, string inputfile, string outputfile)
        {
            CreatePlan cPln = new CreatePlan();
            StreamReader readStream;
            StreamWriter writeStream;
            string[] inputVals;
            List<PlnInfoSim> arcPlnLs = new List<PlnInfoSim>();
            string pt_Ref;
            string crs_Ref;
            string pln_Ref;
            string ds_Ref;
            PlanParameters plnParam_Ref;
            string pt_Tgt;
            string crs_Tgt;
            string pln_Tgt;
            string struct_Tgt;
            PatientSummary ptSumm;
            Patient pt;
            bool cpSucc;
            bool psSucc;

            ////First, create list of patient plans already in the archive////
            readStream = new StreamReader(outputfile);
            //Skip title row//
            readStream.ReadLine();
            while (!readStream.EndOfStream)
            {
                PlnInfoSim plnInfo = new PlnInfoSim();
                inputVals = readStream.ReadLine().Split(',');
                plnInfo.ptId = inputVals[0];
                plnInfo.crsId = inputVals[1];
                plnInfo.plnId = inputVals[2];
                arcPlnLs.Add(plnInfo);
            }
            readStream.Close();
            ////Copy plans from the database file, and paste to the archived patients////
            readStream = new StreamReader(inputfile);
            //Skip title row//
            readStream.ReadLine();
            writeStream = new StreamWriter(outputfile, append: true);
            while (!readStream.EndOfStream)
            {
                ////Read input plan from csv file////
                inputVals = readStream.ReadLine().Split(',');
                pt_Ref = inputVals[0];
                crs_Ref = inputVals[1];
                pln_Ref = inputVals[2];
                ds_Ref = inputVals[3];
                msg = "-Reference:\n  Patient: " + pt_Ref + ", course: " + crs_Ref + ", plan: " + pln_Ref;
                Console.WriteLine(msg);
                //Compare input plan with archived list//
                if (arcPlnLs.Exists(p => p.ptId == pt_Ref && p.crsId == crs_Ref && p.plnId == pln_Ref))
                {
                    Console.WriteLine("--Plan already exists in archive, skipping...\n");
                }
                else
                {
                    ////Copy reference plan parameters////
                    plnParam_Ref = new PlanParameters();
                    cpSucc = plnParam_Ref.extractParametersFromPlan(app, pt_Ref, crs_Ref, pln_Ref);
                    if (cpSucc)
                    {
                        pt_Tgt = "AOA0002";
                        //Check which course to put plan into//
                        ptSumm = app.PatientSummaries.FirstOrDefault(ps => ps.Id == pt_Tgt);
                        pt = app.OpenPatient(ptSumm);
                        crs_Tgt = null;
                        foreach (Course crs in pt.Courses)
                        {
                            //Add at most 10 plans to a course//
                            if (crs.PlanSetups.Count() < 2)
                            {
                                crs_Tgt = crs.Id;
                                break;
                            }
                        }
                        if (crs_Tgt == null)
                        {
                            crs_Tgt = "C" + pt.Courses.Count().ToString();
                        }
                        app.ClosePatient();
                        //Create plan name//
                        pln_Tgt = pt_Ref + crs_Ref.Substring(0, Math.Min(3, crs_Ref.Length)) + pln_Ref.Substring(0, Math.Min(3, pln_Ref.Length));
                        struct_Tgt = "AC_27cm";
                        msg = "-Copy plan to:\n  Patient: " + pt_Tgt + ", course: " + crs_Tgt + ", plan: " + pln_Tgt;
                        Console.WriteLine(msg);
                        ////Copy and create new plan to archived patient////
                        psSucc = cPln.CopyAndCreate(app, pt_Tgt, crs_Tgt, pln_Tgt, struct_Tgt, plnParam_Ref);
                        if (psSucc)
                        {
                            msg = pt_Ref + "," + crs_Ref + "," + pln_Ref + "," + pt_Tgt + "," + crs_Tgt + "," + pln_Tgt + "," + ds_Ref;
                            writeStream.WriteLine(msg);
                        }
                    }
                }
            }
            //Close the file stream//
            readStream.Close();
            writeStream.Close();
        }
        //Simple file info class//
        private class PlnInfoSim
        {
            public string ptId;
            public string crsId;
            public string plnId;
        }
    }
}
