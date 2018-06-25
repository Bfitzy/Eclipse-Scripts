using System;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Runtime.InteropServices;

namespace VMS.TPS
{
  public class Script
  {
    public Script()
    {
    }

        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {
            string finalSummary = "";


            PlanSetup plan = context.PlanSetup;
            Patient patient = context.Patient;
                Structure PTV = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("ptv")).First();
                Structure cord = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("cord")).First();
                Structure cord05 = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("cord +0.5")).Single();
                Structure proxbronch = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("proxbronchtree")).First();
                Structure proxbronch2 = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("proxbronchtree+2")).SingleOrDefault();
                Structure lungR = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("lung r")).First();
            Structure lungL = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("lung l")).First();
                Structure heart = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("heart")).First();
                //Structure ITV = plan.StructureSet.Structures.Where(s=> s.Id.ToLower().Contains("itv")).First();
                //Structure GTV = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("gtvp")).First();
                Structure esophagus = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("esophagus")).First();
                Structure proxtrachea = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("prox trachea")).First();
                Structure aorta = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("aorta")).Single();
                Structure cm2 = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("2cm")).Single();
                Structure chestwall = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("chestwall")).Single();
                //Structure ptvgtv = plan.StructureSet.Structures.Where(s=> s.Id.ToLower().Contains("ptv-gtv do")).Single();
                Structure brachiplex = plan.StructureSet.Structures.Where(s=> s.Id.ToLower().Contains("brachial plexus")).Single();
                Structure skin = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("skin")).Single();
                Structure body = plan.StructureSet.Structures.Where(s => s.Id == "BODY").Single();
                Structure mm2 = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("2mm")).Single();
                Structure lungTotal = plan.StructureSet.Structures.Where(s => s.Id.ToLower().Contains("lung total")).Single();

            DoseValue cordDG = new DoseValue(2250, DoseValue.DoseUnit.cGy);
                DoseValue cordPG = new DoseValue(3000, DoseValue.DoseUnit.cGy);
                DoseValue esophaDG = new DoseValue(2750, DoseValue.DoseUnit.cGy);
                DoseValue esophaPG = new DoseValue(3500, DoseValue.DoseUnit.cGy);
                DoseValue aortaDG = new DoseValue(4700, DoseValue.DoseUnit.cGy);
                DoseValue aortaPG = new DoseValue(1500, DoseValue.DoseUnit.cGy);
                DoseValue brachiDG = new DoseValue(3000, DoseValue.DoseUnit.cGy);
                DoseValue brachiPG = new DoseValue(3200, DoseValue.DoseUnit.cGy);
                DoseValue PTVG = new DoseValue(5000, DoseValue.DoseUnit.cGy);
                DoseValue v50G = new DoseValue(2500, DoseValue.DoseUnit.cGy);
                DoseValue v105 = new DoseValue(5250, DoseValue.DoseUnit.cGy);
                DoseValue v100G = new DoseValue(5000, DoseValue.DoseUnit.cGy);
                DoseValue v20G = new DoseValue(2000, DoseValue.DoseUnit.cGy);
                DoseValue tracheaDG = new DoseValue(1800, DoseValue.DoseUnit.cGy);
                DoseValue stomachDG = new DoseValue(2800, DoseValue.DoseUnit.cGy);
                DoseValue lung15 = new DoseValue(1250, DoseValue.DoseUnit.cGy);
                DoseValue lung10 = new DoseValue(1350, DoseValue.DoseUnit.cGy);
                DoseValue chestVD = new DoseValue(3000, DoseValue.DoseUnit.cGy);
                DoseValue v90 = new DoseValue(4500, DoseValue.DoseUnit.cGy);

            double cordVG = plan.GetVolumeAtDose(cord, cordDG, VolumePresentation.AbsoluteCm3);
                double PTV_vol = PTV.Volume;
                double D50_vol = plan.GetVolumeAtDose(body, v50G, VolumePresentation.AbsoluteCm3);
                double D100_vol = plan.GetVolumeAtDose(body, v100G, VolumePresentation.AbsoluteCm3);
                double V20 = plan.GetVolumeAtDose(lungTotal, v20G, VolumePresentation.Relative);

                double esophagV = plan.GetVolumeAtDose(esophagus, esophaDG, VolumePresentation.AbsoluteCm3);
                double brachiV = plan.GetVolumeAtDose(brachiplex, brachiDG, VolumePresentation.AbsoluteCm3);
                double heartV = plan.GetVolumeAtDose(heart, brachiPG, VolumePresentation.AbsoluteCm3);
                double tracheaV = plan.GetVolumeAtDose(proxtrachea, tracheaDG, VolumePresentation.AbsoluteCm3);
                double skinV = plan.GetVolumeAtDose(skin, cordPG, VolumePresentation.AbsoluteCm3);
                //double stomachV = plan.GetVolumeAtDose(stomach, stomachDG, VolumePresentation.AbsoluteCm3);
                double chestV = plan.GetVolumeAtDose(chestwall, cordPG, VolumePresentation.AbsoluteCm3);
                double lung15RV = plan.GetVolumeAtDose(lungR, lung15, VolumePresentation.AbsoluteCm3);
                double lung15LV = plan.GetVolumeAtDose(lungL, lung15, VolumePresentation.AbsoluteCm3);
                double lung10RV = plan.GetVolumeAtDose(lungR, lung10, VolumePresentation.AbsoluteCm3);
                double lung10LV = plan.GetVolumeAtDose(lungL, lung10, VolumePresentation.AbsoluteCm3);
            double cov100 = plan.GetVolumeAtDose(PTV, v100G, VolumePresentation.Relative);
            double cov90 = plan.GetVolumeAtDose(PTV, v90, VolumePresentation.Relative);
            double aortaV = plan.GetVolumeAtDose(aorta, aortaDG, VolumePresentation.AbsoluteCm3);
            double cm2v = plan.GetVolumeAtDose(cm2, v105, VolumePresentation.AbsoluteCm3);
            double bronchV = plan.GetVolumeAtDose(proxbronch, esophaPG, VolumePresentation.AbsoluteCm3);


            DoseValue ptvmax = plan.GetDVHCumulativeData(PTV, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
            DoseValue cm2P = plan.GetDVHCumulativeData(cm2, DoseValuePresentation.Relative, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
            DoseValue esophagP = plan.GetDVHCumulativeData(esophagus, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
            DoseValue cordP = plan.GetDVHCumulativeData(cord, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
            DoseValue mm2P = plan.GetDVHCumulativeData(mm2, DoseValuePresentation.Relative, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
            DoseValue aortaP = plan.GetDVHCumulativeData(aorta, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
                DoseValue brachiP = plan.GetDVHCumulativeData(brachiplex, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
                DoseValue heartP = plan.GetDVHCumulativeData(heart, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
                DoseValue tracheaP = plan.GetDVHCumulativeData(proxtrachea, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
                DoseValue skinP = plan.GetDVHCumulativeData(skin, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
                //DoseValue stomachP = plan.GetDVHCumulativeData(stomach, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
                DoseValue chestP = plan.GetDVHCumulativeData(chestwall, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;
            DoseValue bronchP = plan.GetDVHCumulativeData(proxbronch, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1).MaxDose;


            double hetero = 500000 / ptvmax.Dose;
            finalSummary += "Objective,Actual,Pass-Fail\nCoverage\n";
            finalSummary += "95% of PTV covered by 100% of Pres Dose," + Math.Round(cov100, 2) + "%," + PWF(20, 5, 100 - cov100) + '\n';
            finalSummary += "99% of PTV receives minimum of 90% prescribed dose," + Math.Round(cov90, 2) + "%," + PWF(5, 1, 100 - cov90) + '\n';
            finalSummary += "Dose Heterogeneity - Target Dose > 60% 0f PTV Max," + Math.Round(hetero, 2) + "%," + PWF(60, 40, 100 - hetero) + '\n';
            finalSummary += "High Dose Spillage - No Dose Greater than 105% outside PTV," + cm2P.ToString() + ',' + PWF(130, 105, cm2P.Dose) + '\n';
            finalSummary += "Spillage Volume:," + Math.Round(cm2v, 2) + " cc," + PWF(10, 1, cm2v) + '\n';
            finalSummary += "PTV Volume:," + Math.Round(PTV_vol, 2) + " cc" + '\n';
            finalSummary += "V100 / PTV:," + Math.Round(D100_vol / PTV_vol, 2) + ',' + PWF(1.5, 1.2, D100_vol / PTV_vol) + '\n';
            finalSummary += "V50 Volume:," + Math.Round(D50_vol, 2) + " cc\n";
            finalSummary += "D50_volume / PTV volume:," + Math.Round(D50_vol / PTV_vol, 2) + ',' + PWF(-.743*Math.Log(PTV_vol)+7.7335, -.668*Math.Log(PTV_vol)+6.4269, D50_vol / PTV_vol) + '\n';

            finalSummary += "2cm max dose:," + mm2P.ToString() + ',' + PWF(-4E-9*Math.Pow(PTV_vol,5)+2E-6*Math.Pow(PTV_vol, 4)-4E-4*Math.Pow(PTV_vol,3)+265E-4*Math.Pow(PTV_vol,2)-2045E-4*PTV_vol + 57.527, -4E-9 * Math.Pow(PTV_vol,5) + 2E-6*Math.Pow(PTV_vol,4) - 2E-4*Math.Pow(PTV_vol,3)+ 149E-4* Math.Pow(PTV_vol,2) -.0578 * PTV_vol + 49.81, mm2P.Dose) + '\n';
            finalSummary += "Critical Structures\n";
            finalSummary += "Lung V20: (percent of Volume <20%)," + Math.Round(V20, 2) + '%' + ',' + PWF(15, 10, V20) + '\n';
            finalSummary += "Spinal Cord Point Max: (< 30 Gy)," + cordP.ToString() + ',' + PWF(3000, 2800, cordP.Dose) + '\n';
            finalSummary += "Spinal Cord V@22.5 Gy: (<.25cc)," + Math.Round(cordVG, 5) + " cc," + PWF(.25,.1, cordVG) + "\n";

            finalSummary += "Esophagus Max Point Dose: (<35Gy)," + esophagP.ToString() + ',' + PWF(3500, 3000, esophagP.Dose) + '\n';
            finalSummary += "Esophagus V@ 27.5 Gy: (<5cc)," + Math.Round(esophagV, 2) + "cc," + PWF(5, 4, esophagV) + '\n';
            finalSummary += "Brachial Plexus max point dose: (<32 Gy)," + brachiP.ToString() + ',' + PWF(3200, 3000, brachiP.Dose) + '\n';
            finalSummary += "Brachial Plexus volume @30Gy: (<3cc)," + Math.Round(brachiV, 2) + " cc," + PWF(3, 2.5, brachiV) + "\n";

            finalSummary += "Heart Point Max: (<38 Gy)," + heartP.ToString() + ',' + PWF(3800, 3500, heartP.Dose) + '\n';
            finalSummary += "Heart V@32 Gy: (<3cc)," + Math.Round(heartV, 2) + " cc,"+ PWF(15,10,heartV) +"\n";

            finalSummary += "Aorta max point dose: (<53 Gy)," + aortaP.ToString() + ',' + PWF(5300, 5100, aortaP.Dose) + '\n';
            finalSummary += "Aorta V@ 47Gy: (<10cc)," + Math.Round(aortaV, 2) + ',' + PWF(10, 8, aortaV) + '\n';
            finalSummary += "Trachea Point Max: (<38 Gy)," + tracheaP.ToString() + ',' + PWF(3800, 3500, tracheaP.Dose) + '\n';
            finalSummary += "Trachea V@18 Gy: (<4cc)," + Math.Round(tracheaV, 2) + " cc," + PWF(4, 3, tracheaV) + "\n";
            finalSummary += "Bronchus Point Max: (<38 Gy)," + bronchP.ToString() + ',' + PWF(3800, 3500, bronchP.Dose) + '\n';
            finalSummary += "Bronchus V@18 Gy: (< 4cc), " + Math.Round(bronchV, 2) + " cc," + PWF(4, 3, bronchV) + "\n";
            finalSummary += "Skin Point Max: (<32 Gy)," + skinP.ToString() + ',' + PWF(3200, 3000, skinP.Dose) + '\n';
            finalSummary += "Skin V@30 Gy: (<30cc)," + Math.Round(skinV, 2) + " cc,"+ PWF(10, 8, skinV) +"\n";

            //finalSummary += "Stomach V@28 Gy: " + Math.Round(stomachV, 2) + " cc,"+ PWF(10, 8, stomachV) +"\n";
            //finalSummary += "Stomach Point Max: " + stomachP.ToString() +','+ PWF(3200, 3000) +'\n';

            //finalSummary += "Chest Wall Point Max: ()," + chestP.ToString() + ',' + '\n';
            finalSummary += "Chest Wall V@30 Gy: (<30cc)," + Math.Round(chestV, 2) + " cc,"+ PWF(30, 25, chestV) +"\n";

                finalSummary += "Right Lung V@12.5 Gy: (<1500cc)," + Math.Round(lung15RV, 2) + " cc,"+ PWF(1500, 1250, lung15RV) +"\n";
                finalSummary += "Left Lung V@12.5 Gy: (<1500cc)," + Math.Round(lung15LV, 2) + " cc," + PWF(1500, 1250, lung15LV) + "\n";
                finalSummary += "Right Lung V@13.5 Gy: (<1000cc)," + Math.Round(lung10RV, 2) + " cc,"+ PWF(1000, 800, lung10RV) + "\n";
                finalSummary += "Left Lung V@13.5 Gy: (<1000cc)," + Math.Round(lung10LV, 2) + " cc,"+ PWF(1000, 800, lung10LV) + "\n";
                StreamWriter Sw = new StreamWriter("v:\\Special_Phys_Reports\\"+patient.LastName + '_'+patient.Id+"Lung_5x10.csv");
                Sw.WriteLine(finalSummary);
                Sw.Close();



        }

    public string PWF(double fail, double warn, double value)
        {
            if(value < warn)
            {
                return "Pass";
            } else if((value < fail)||(value >=warn))
            {
                return "Warning";
            } else
            {
                return "FAIL";
            }
        }
  }
}
