//////////////////////////////////////////////////////////////////////
///Copy plan parameters from the reference plan and create a plan based on the parameters
///version 0.1
///add Calculation Model
///version 0.0
///Becket Hui 2020/9
//////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Configuration;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace CreatePatientDataBase
{
    public class CreatePlan
    {
        public bool CopyAndCreate(Application app, string ptId_Tgt, string crsId_Tgt, string plnId_Tgt, string structId_Tgt, PlanParameters plnParam_Ref)
        {
            //Open patient//
            PatientSummary ptSumm = app.PatientSummaries.FirstOrDefault(ps => ps.Id == ptId_Tgt);
            if (ptSumm == null)
            {
                Console.WriteLine("--Cannot find patient" + ptId_Tgt + ".");
                return false;
            }
            Patient pt = app.OpenPatient(ptSumm);
            try
            {
                pt.BeginModifications();
                //Open or create course//
                Course crs = pt.Courses.FirstOrDefault(c => c.Id == crsId_Tgt);
                if (crs == null)
                {
                    Console.WriteLine("-Create course " + crsId_Tgt + ".");
                    crs = pt.AddCourse();
                    crs.Id = crsId_Tgt;
                }
                //Create plan//
                ExternalPlanSetup pln = crs.ExternalPlanSetups.FirstOrDefault(p => p.Id == plnId_Tgt);
                if (pln == null)
                {
                    StructureSet structSet = pt.StructureSets.FirstOrDefault(ss => ss.Id == structId_Tgt);
                    if (structSet == null)
                    {
                        Console.WriteLine("--Cannot find structure set " + structId_Tgt + ". Plan is not created\n");
                        app.ClosePatient();
                        return false;
                    }
                    pln = crs.AddExternalPlanSetup(structSet);
                    pln.Id = plnId_Tgt;
                }
                //Return if there is already a plan with the same name//
                else
                {
                    Console.WriteLine("--A plan with name " + plnId_Tgt + " already exists. Plan is not created\n");
                    app.ClosePatient();
                    return false;
                }
                Console.WriteLine("-Start creating plan " + plnId_Tgt + ".");
                //Set plan prescription properties//
                pln.SetPrescription(plnParam_Ref.N_Fx, plnParam_Ref.DoseperFx, plnParam_Ref.TrtPct);
                ///////////Create beam by copying from the beams in reference plan parameters////////////
                //Create empty list of MU values for each beam//
                List<KeyValuePair<string, MetersetValue>> muValues = new List<KeyValuePair<string, MetersetValue>>();
                foreach (PlanParameters.BmParam bmParam in plnParam_Ref.BmParamLs)
                {
                    //Add beam, type based on reference MLC beam technique//
                    IEnumerable<double> muSet = bmParam.CtrPtParam.ControlPoints.Select(cp => cp.MetersetWeight).ToList();
                    switch (bmParam.MLCBmTechnique)
                    {
                        case "StaticMLC":
                            Beam bm =
                                pln.AddMLCBeam(bmParam.MachParam, new float[2, 60],
                                new VRect<double>(-10.0, -10.0, 10.0, 10.0), 0.0, 0.0, 0.0,bmParam.CtrPtParam.Isocenter);
                            bm.Id = bmParam.bmId;
                            bm.ApplyParameters(bmParam.CtrPtParam);
                            break;
                        case "StaticSegWin":
                            bm =
                                pln.AddMultipleStaticSegmentBeam(bmParam.MachParam, muSet, 0.0, 0.0, 0.0, bmParam.CtrPtParam.Isocenter);
                            bm.Id = bmParam.bmId;
                            muValues.Add(new KeyValuePair<string, MetersetValue>(bmParam.bmId, bmParam.mu));
                            bm.ApplyParameters(bmParam.CtrPtParam);
                            break;
                        case "StaticSlidingWin":
                            bm =
                                pln.AddSlidingWindowBeam(bmParam.MachParam, muSet, 0.0, 0.0, 0.0, bmParam.CtrPtParam.Isocenter);
                            bm.Id = bmParam.bmId;
                            muValues.Add(new KeyValuePair<string, MetersetValue>(bmParam.bmId, bmParam.mu));
                            bm.ApplyParameters(bmParam.CtrPtParam);
                            break;
                        case "ConformalArc":
                            bm =
                                pln.AddConformalArcBeam(bmParam.MachParam, 0.0, bmParam.CtrPtParam.ControlPoints.Count(),
                                bmParam.CtrPtParam.ControlPoints.First().GantryAngle, bmParam.CtrPtParam.ControlPoints.Last().GantryAngle,
                                bmParam.CtrPtParam.GantryDirection, 0.0, bmParam.CtrPtParam.Isocenter);
                            bm.Id = bmParam.bmId;
                            bm.ApplyParameters(bmParam.CtrPtParam);
                            break;
                        case "VMAT":
                            bm =
                                pln.AddVMATBeam(bmParam.MachParam, muSet, 0.0,
                                bmParam.CtrPtParam.ControlPoints.First().GantryAngle, bmParam.CtrPtParam.ControlPoints.Last().GantryAngle,
                                bmParam.CtrPtParam.GantryDirection, 0.0, bmParam.CtrPtParam.Isocenter);
                            bm.Id = bmParam.bmId;
                            bm.ApplyParameters(bmParam.CtrPtParam);
                            break;
                        default:
                            Console.WriteLine("--At least one of the beams is unidentified, plan is not created.\n");
                            app.ClosePatient();
                            return false;
                    }
                }
                //Set the plan normalization value//
                pln.PlanNormalizationValue = plnParam_Ref.PlnNormFactr;
                //Set the plan calculation model//
                pln.SetCalculationModel(CalculationType.PhotonVolumeDose, plnParam_Ref.CalcModel);
                //If one of the beams is static IMRT, compute dose to enforce MUs to beams//
                if (plnParam_Ref.BmParamLs.Any(bm => bm.MLCBmTechnique == "StaticSegWin" || bm.MLCBmTechnique == "StaticSlidingWin"))
                {
                    Console.WriteLine("--Start computing static beam IMRT plan.");
                    pln.CalculateDoseWithPresetValues(muValues);
                }
                Console.WriteLine("-Finish plan creation.\n");
                app.SaveModifications();
                app.ClosePatient();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                app.ClosePatient();
                return false;
            }
        }
    }
}
