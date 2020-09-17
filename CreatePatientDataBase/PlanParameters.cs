//////////////////////////////////////////////////////////////////////
///Class that declares properties of the reference plan and method to read the properties from the reference plan
///version 0.1
///add Calculation Model
///version 0.0
///Becket Hui 2020/9
//////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace CreatePatientDataBase
{
    public class PlanParameters
    {
        public string PtId { get; set; }
        public string CrsId { get; set; }
        public string PlnId { get; set; }
        public DateTime ModDate { get; set; }
        public int N_Fx { get; set; }
        public DoseValue DoseperFx { get; set; }
        public double TrtPct { get; set; }
        public double PlnNormFactr { get; private set; }
        public string CalcModel { get; set; }
        public List<BmParam> BmParamLs { get; private set; }
        public class BmParam
        {
            public string bmId;
            public string machId;
            public string energy;
            public int doseRt;
            public string bmTechnique;
            public MLCPlanType mlcTyp;
            public ExternalBeamMachineParameters MachParam;
            public BeamParameters CtrPtParam;
            public string MLCBmTechnique;
            public MetersetValue mu;
        }
        //Read and assign parameters from input patient, course and plan//
        public bool extractParametersFromPlan(Application app, string ptId_Input, string crsId_Input, string plnId_Input)
        {
            PtId = ptId_Input;
            CrsId = crsId_Input;
            PlnId = plnId_Input;
            BmParamLs = new List<BmParam>();
            try
            {
                //Read plan//
                PatientSummary ptSumm = app.PatientSummaries.FirstOrDefault(ps => ps.Id == PtId);
                if (ptSumm == null)
                {
                    Console.WriteLine("--Cannot find patient.\n");
                    return false;
                }
                Patient pt = app.OpenPatient(ptSumm);
                Course crs = pt.Courses.FirstOrDefault(c => c.Id == CrsId);
                if (crs == null)
                {
                    Console.WriteLine("--Cannot find course.\n");
                    return false;
                }
                ExternalPlanSetup pln = crs.ExternalPlanSetups.FirstOrDefault(p => p.Id == PlnId);
                if(pln == null)
                {
                    Console.WriteLine("--Cannot find plan.\n");
                    return false; 
                }
                //get last modified date time//
                ModDate = pln.HistoryDateTime;
                //get prescription info//
                N_Fx = Math.Max(pln.NumberOfFractions.GetValueOrDefault(), 1); //At least 1 fx is needed
                DoseperFx = pln.DosePerFraction;
                TrtPct = pln.TreatmentPercentage;
                //get plan normalization factor//
                PlnNormFactr = pln.PlanNormalizationValue;
                //get calculation model//
                CalcModel = pln.PhotonCalculationModel;
                //get beam parameters//
                foreach (Beam bm in pln.Beams)
                {
                    if (bm.MetersetPerGy > 0)
                    {
                        BmParam bp = new BmParam
                        {
                            bmId = bm.Id,
                            machId = bm.TreatmentUnit.Id.ToString(),
                            energy = bm.EnergyModeDisplayName,
                            doseRt = bm.DoseRate,
                            bmTechnique = bm.Technique.Id.ToString(),
                            mlcTyp = bm.MLCPlanType,
                            CtrPtParam = bm.GetEditableParameters(),
                            mu = bm.Meterset
                        };
                        bp.MachParam = 
                            new ExternalBeamMachineParameters(bp.machId, bp.energy, bp.doseRt, bp.bmTechnique, null);
                        bp.MLCBmTechnique = setMLCBmTechnique(bp, bm.CalculationLogs);
                        if (bp.MLCBmTechnique == null)
                        {
                            Console.WriteLine("--At least one of the beams is invalid, skipping...\n");
                            return false;
                        }
                        BmParamLs.Add(bp);
                    }
                }
                app.ClosePatient();
                Console.WriteLine("-Finish reading parameters from reference.");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                app.ClosePatient();
                return false;
            }
        }
        //Determine the add beam method based on beam technique and mlc technique//
        private string setMLCBmTechnique(BmParam bp, IEnumerable<BeamCalculationLog> calcLog)
        {
            if (bp.bmTechnique == "STATIC" && bp.mlcTyp == MLCPlanType.Static)
            {
                return "StaticMLC";
            }
            else if (bp.bmTechnique == "STATIC" && bp.mlcTyp == MLCPlanType.DoseDynamic)
            {
                // Check if the MLC technique is Sliding Window or Segmental
                var lines = calcLog.FirstOrDefault(log => log.Category == "LMC");
                foreach (var line in lines.MessageLines)
                {
                    if (line.ToUpper().Contains("MULTIPLE STATIC SEGMENTS")) { return "StaticSegWin"; }
                    if (line.ToUpper().Contains("SLIDING WINDOW")) { return "StaticSlidingWin"; }
                    if (line.ToUpper().Contains("SLIDING-WINDOW")) { return "StaticSlidingWin"; }
                }
                return null;
            }
            else if (bp.bmTechnique == "ARC" && bp.mlcTyp == MLCPlanType.ArcDynamic)
            {
                return "ConformalArc";
            }
            else if (bp.bmTechnique == "ARC" && bp.mlcTyp == MLCPlanType.VMAT)
            {
                return "VMAT";
            }
            else
            {
                return null;
            }
        }
    }
}
