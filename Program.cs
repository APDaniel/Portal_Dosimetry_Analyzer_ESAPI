using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using VMS.CA.Scripting;
using VMS.DV.PD.Scripting;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace PDAnalyzer
{

using System.Windows;
using System.Data;
    static class Program
    {

        [STAThread]
        static void Main(string[] args)
        {
            
            try
            {
                using (VMS.DV.PD.Scripting.Application application = VMS.DV.PD.Scripting.Application.CreateApplication())
                {
                    Execute(application);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                Console.ReadLine();
            }
        }


        static void Execute(VMS.DV.PD.Scripting.Application application)
        {
            //===========================================================================================================================================================================
            //===========================================================================================================================================================================
            //for licensing purpose
            var current_date = DateTime.Now;
            
            DateTime expiration_date = new DateTime(2021, 12, 1);
            DateTime warning_date = new DateTime(2021, 11, 23);
            string user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            if (current_date>expiration_date)
            {
                Console.WriteLine("Dear {0},\n Thank you for your interest in using this script.\n Unfortunately, this version is expired.\n\n If you would like to proceed, contact to Daniel Peidus ('daniel.peudys@gmail.com' or 'd.peidus@pet-net.ru')", user);
                return;
            }
            if (current_date>warning_date)
            {
                Console.WriteLine("Dear {0},\n This message is to warn you that this version of software will expire soon. The expiration date is 12/01/2021.\n\n To purchase the next version, contact to Daniel Peidus ('daniel.peudys@gmail.com' or 'd.peidus@pet-net.ru')\n\n\nIf you would like to proceed with execution,\nplease, press 'ENTER'\n", user);
            } 
            //===========================================================================================================================================================================
            //===========================================================================================================================================================================

            VMS.DV.PD.UI.Base.VTransientImageDataMgr.CreateInstance(true);
            
            //Start stopwatch
            DateTime stopwatch_start = DateTime.Now;
            DateTime start_date = DateTime.Today;

            //Predefine analysis parameters by default

            var GammaParameterDoseDifference = 0.03;
            var GammaParameterDistanceToAgreement = 3;
            
            int specific_patient = 0;
            string patientID = "";

            string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            Console.WriteLine("Greetings, {0}!", username);
            Console.WriteLine("");
            
            Console.WriteLine("Please, press any key to set custom analysis date\nAfter the 60 seconds, the application will perform analysis for today");

            
            //Timer for user's inactivity. If no key will be pressed, analyzis will be performed for today
            DateTime timer = DateTime.Now;
            var stop_timer = Math.Round((((DateTime.Now - timer).TotalSeconds)));

            
                for (stop_timer = 0; stop_timer < 60; stop_timer = Math.Round((((DateTime.Now - timer).TotalSeconds))))
                {

                    Console.SetCursorPosition(0, 5);
                    Console.CursorVisible = false;
                    stop_timer = Math.Round((((DateTime.Now - timer).TotalSeconds)));
                    Console.Write("          " + stop_timer);
                    Console.Write("\b\b\b\b\b ");
                    if (Console.KeyAvailable)
                {
                    definedate:
                    try
                    {
                        
                        Console.WriteLine("\nPlease, enter YEAR for start date of analysis");
                        int year = Convert.ToInt32(Console.ReadLine());

                        Console.WriteLine("\nPlease, enter MONTH for start date of analysis");
                        int month = Convert.ToInt32(Console.ReadLine());

                        Console.WriteLine("\nPlease, enter DAY for start date of analysis");
                        int day = Convert.ToInt32(Console.ReadLine());

                        start_date = new DateTime(year, month, day);
                        goto patient_specification;
                    }
                    catch
                    {
                        Console.Clear();
                        Console.Write("Unfortunately, the format of the date was not recognized...\nDo you wish to try again? (Y/N)");

                        if (Console.ReadKey().Key == ConsoleKey.Y) {Console.Clear(); goto definedate; }
                        if (Console.ReadKey().Key == ConsoleKey.N) {Console.Clear();start_date = DateTime.Today; goto patient_specification; }
                        else { Console.Clear(); start_date = DateTime.Today; goto patient_specification; }
                        
                    }
                }
                    if (stop_timer>58)
                {
                    specific_patient = 1;
                    goto analysis;
                }
                }


            




        patient_specification:

        Console.CursorVisible = true;
        Console.WriteLine("\nDo you wish to analyze specific patient? Y/N");
        Console.WriteLine("Please, enter 'yy' for specifying certaing patient or 'nn' for analyzing through all the database");
        if (Console.ReadKey().Key==ConsoleKey.Y) {specific_patient = 2; goto analysis;}
        if (Console.ReadKey().Key==ConsoleKey.N) {specific_patient = 1; goto analysis;}
        
        



            analysis:
            
            Console.WriteLine("\n\n");
            Console.SetCursorPosition(0, 15);
            Console.WriteLine("\n\n\n\nStarting analysis for {0}...\n", start_date);
            if (specific_patient > 1)
            {
                tapPatientsID:
                Console.WriteLine("Please, enter patient's ID");
                Console.WriteLine("");
                try { patientID = Convert.ToString(Console.ReadLine()); }
                catch 
                { 
                    Console.WriteLine("The patient's ID is incorrect.\nIf you would like to try other ID, press 'Y'\nTo perform analysis through all database, press 'N'\n");
                    if (Console.ReadKey().Key == ConsoleKey.Y) { Console.Clear(); goto tapPatientsID; }
                    if (Console.ReadKey().Key == ConsoleKey.N) { Console.Clear(); specific_patient = 1; goto analysis; }
                }
            }


            Console.SetCursorPosition(0, 29);
            Console.WriteLine("\n\n\nDo you wish to adjust Gamma parameters? Defaults are: DTA=3mm, DoseDiff=3% (Y/N)\nAfter the 60 seconds setting by default will be applied");
            

            timer = DateTime.Now;
            stop_timer = Math.Round((((DateTime.Now - timer).TotalSeconds)));


            //Define GammaParameters if needed
            for (stop_timer = 0; stop_timer < 60; stop_timer = Math.Round((((DateTime.Now - timer).TotalSeconds))))
            {

                Console.SetCursorPosition(0, 35);
                Console.CursorVisible = false;
                stop_timer = Math.Round((((DateTime.Now - timer).TotalSeconds)));
                Console.Write("          " + stop_timer);
                Console.Write("\n\n\n");
                
                if (Console.KeyAvailable)
                {
                    if (Console.ReadKey().Key == ConsoleKey.Y) 
                    
                    {
                        GAMMAadjustment:
                        Console.WriteLine("Please, enter DTA in mm:\n");
                        try { GammaParameterDistanceToAgreement = Convert.ToInt32(Console.ReadLine()); }
                        catch { Console.WriteLine("Seems, that the format of the entered number is not correct...\nPlease, enter DTA in mm"); goto GAMMAadjustment; }
                        Console.Clear();
                        Console.WriteLine("Please, enter DoseDifference in precents (%):\n");
                        try { GammaParameterDoseDifference =(Convert.ToDouble(Console.ReadLine()) / 100); goto Start_analysis; }
                        catch { Console.WriteLine("Seems, that the format of the entered number is not correct...\nPlease, enter DoseDiff in percents(%)"); goto GAMMAadjustment; }
                    }
                    if (Console.ReadKey().Key == ConsoleKey.N) { goto Start_analysis; }
                }
                
                if (stop_timer > 58)
                {
                    
                    goto Start_analysis;
                }
            }


        

        Start_analysis:
            Console.WriteLine("\n\n\nAnalysis for {0} started...", start_date);
            //Turn on beams counter for a loop
            int iterator = 0;
            int number_of_analyzed_objects = 0;
            int counter_is_the_patient_is_counted = 0;
            int number_of_patients_database = 0;
            int number_of_patients_analyzed = 0;


            //Define metrics to fill columns in an excel report file
            List<string> ID_List = new List<string>();
            List<string> FN_List = new List<string>();
            List<string> LN_List = new List<string>();
            List<string> Course_List = new List<string>();
            List<string> PlanID_List = new List<string>();
            List<string> BeamID_List = new List<string>();
            List<DateTime> TreatedDate_List = new List<DateTime>();
            List<double> MDD_List = new List<double>();
            List<double> ADD_List = new List<double>();
            List<double> MG_List = new List<double>();
            List<double> GLTO_List = new List<double>();

            if (specific_patient > 1)
            {
                try
                {
                    application.ClosePatient();
                    Patient test_patient = application.OpenPatientById(patientID);
                }
                catch
                {
                    Console.WriteLine("I very much regret, but ID seems to be incorrect\n");
                    Console.WriteLine("If you would like to try other ID, press 'Y'\nTo perform analysis through all database, press 'N'\n");
                    if (Console.ReadKey().Key == ConsoleKey.Y) { Console.Clear(); goto _tapPatientsID; }
                    if (Console.ReadKey().Key == ConsoleKey.N) { Console.Clear(); specific_patient = 0; goto analysis; }
                _tapPatientsID:
                    //patientID = Convert.ToString(Console.ReadLine());
                    try { patientID = Convert.ToString(Console.ReadLine()); }
                    catch
                    {
                        Console.WriteLine("The patient's ID is incorrect.\nIf you would like to try other ID, press 'Y'\nTo perform analysis through all database, press 'N'\n");
                        if (Console.ReadKey().Key == ConsoleKey.Y) { Console.Clear(); goto _tapPatientsID; }
                        if (Console.ReadKey().Key == ConsoleKey.N) { Console.Clear(); specific_patient = 0; goto analysis; }
                    }
                }
                application.ClosePatient();
                Patient patient = application.OpenPatientById(patientID);
                if (patient != null)
                {


                    foreach (PDPlanSetup pd_plan in patient.PDPlanSetups.OrderByDescending(z => z.HistoryDateTime))
                    {


                        foreach (PDBeam pd_beam in pd_plan.Beams.OrderByDescending(z => z.HistoryDateTime))


                        //Where(x=>x.Beam.CreationDateTime>start_date)
                        {
                            iterator++;
                            number_of_analyzed_objects++;
                            if (iterator > 10000000)

                            {
                                Console.WriteLine("\n\n\nThis program was made an attempt to analyze more than 1 million objects.\nExecution is interrupted for safety purposes!\nTo exit the application, please, press ENTER...");
                                Console.ReadLine();
                                goto Finished;
                            }

                            //define metrics to show to a user
                            double ResultmaxdoseDiffRelative = 0;
                            double ResultAverageDD = 0;
                            double ResultMaxGamma = 0;
                            double ResultGammaLessThanOne = 0;
                            try
                            {


                                List<EvaluationTestDesc> evaluationTestDescs = new List<EvaluationTestDesc>();
                                EvaluationTestDesc Desc_MaxDose = new EvaluationTestDesc(EvaluationTestKind.MaxDoseDifferenceRelative, double.NaN, 1, false);
                                EvaluationTestDesc Desc_AverageDD = new EvaluationTestDesc(EvaluationTestKind.AverageDoseDifferenceRelative, double.NaN, 3, false);
                                EvaluationTestDesc Desc_MaxGamma = new EvaluationTestDesc(EvaluationTestKind.MaxGamma, double.NaN, 3, false);
                                EvaluationTestDesc Desc_GammaLessThanOne = new EvaluationTestDesc(EvaluationTestKind.GammaAreaLessThanOne, double.NaN, 100, false);
                                evaluationTestDescs.Add(Desc_MaxDose);
                                evaluationTestDescs.Add(Desc_AverageDD);
                                evaluationTestDescs.Add(Desc_MaxGamma);
                                evaluationTestDescs.Add(Desc_GammaLessThanOne);
                                //evaluationTestDescs.Add(Desk_GammaLessThanOne);

                                PDTemplate templatePD = new PDTemplate(false, false, false, false, AnalysisMode.Relative, NormalizationMethod.MinimizeDifference, true, 0.2, ROIType.CIAO, 5, GammaParameterDoseDifference, GammaParameterDistanceToAgreement, false, evaluationTestDescs);


                                //PortalDoseImage reference_image = pd_beam.PortalDoseImages.Last();
                                PortalDoseImage image_to_analyze = pd_beam.PortalDoseImages.Last();
                                DoseImage reference_image = pd_beam.PortalDoseImages.First();
                                DoseImage baseline_image = pd_beam.ConstancyCheckBaselineImage;

                                if (baseline_image != null)
                                {
                                    reference_image = baseline_image;
                                }

                                PDAnalysis analysis = new PDAnalysis();



                                //Lets use try/catch to avoid malfunctions
                                foreach (PortalDoseImage PDimage in pd_beam.PortalDoseImages.OrderByDescending(z=>z.HistoryDateTime))
                                {
                                    if (PDimage.Image.CreationDateTime > start_date)
                                    {
                                        try
                                        {

                                            analysis = PDimage.CreateTransientAnalysis(templatePD, reference_image);

                                            EvaluationTest TestmaxdoseDiffRelative = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifferenceRelative);
                                            ResultmaxdoseDiffRelative = Math.Round(TestmaxdoseDiffRelative.TestValue * 100, 2);


                                            EvaluationTest TestAverageDD = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifferenceRelative);
                                            ResultAverageDD = Math.Round(TestAverageDD.TestValue * 100, 2);

                                            EvaluationTest TestMaxGamma = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxGamma);
                                            ResultMaxGamma = Math.Round(TestMaxGamma.TestValue, 2);

                                            EvaluationTest TestGammaLessThanOne = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.GammaAreaLessThanOne);
                                            ResultGammaLessThanOne = Math.Round(TestGammaLessThanOne.TestValue * 100, 2);

                                            //Fulfill lists with the data to be exported to excel
                                            ID_List.Add(patient.Id);
                                            FN_List.Add(patient.FirstName);
                                            LN_List.Add(patient.LastName);
                                            Course_List.Add(pd_plan.PlanSetup.Course.Id);
                                            PlanID_List.Add(pd_plan.Id);
                                            BeamID_List.Add(pd_beam.Id);
                                            TreatedDate_List.Add(Convert.ToDateTime(PDimage.Image.CreationDateTime));
                                            MDD_List.Add(ResultmaxdoseDiffRelative);
                                            ADD_List.Add(ResultAverageDD);
                                            MG_List.Add(ResultMaxGamma);
                                            GLTO_List.Add(ResultGammaLessThanOne);

                                            if (ResultGammaLessThanOne < 90 && ResultMaxGamma > 2)
                                            {
                                                Console.BackgroundColor = default;
                                                Console.BackgroundColor = ConsoleColor.DarkRed;
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                Console.BackgroundColor = default;
                                            }
                                            else if ((pd_plan.Id.Contains("H")&&pd_plan.Id.Contains("N")) && (ResultmaxdoseDiffRelative > 10 || ResultAverageDD > 1 || ResultMaxGamma > 2))
                                            {
                                                Console.BackgroundColor = default;
                                                Console.BackgroundColor = ConsoleColor.DarkMagenta;
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                Console.BackgroundColor = default;
                                            }
                                            else if (pd_plan.Id.Contains("Br") && (ResultmaxdoseDiffRelative > 10 || ResultAverageDD > 1 || ResultMaxGamma > 2))
                                            {
                                                Console.BackgroundColor = default;
                                                Console.BackgroundColor = ConsoleColor.DarkMagenta;
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                Console.BackgroundColor = default;
                                            }
                                            else if (ResultmaxdoseDiffRelative <1 && ResultAverageDD < 1 && ResultMaxGamma < 1)
                                            {
                                                Console.BackgroundColor = default;
                                                Console.BackgroundColor = ConsoleColor.DarkGray;
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, Probably, it is a portal QA field or reference dose image");
                                                Console.BackgroundColor = default;
                                            }
                                            else if (ResultmaxdoseDiffRelative>80)
                                            {
                                                Console.BackgroundColor = default;
                                                Console.BackgroundColor = ConsoleColor.DarkCyan;
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                Console.BackgroundColor = default;
                                            }
                                            else
                                            {
                                                Console.BackgroundColor = default;
                                                Console.ForegroundColor = ConsoleColor.Gray;
                                                Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                Console.BackgroundColor = default;
                                            }
                                            counter_is_the_patient_is_counted = 1;
                                        }
                                        catch
                                        {
                                            Console.ForegroundColor = ConsoleColor.DarkGray;
                                            Console.WriteLine($"No first-fraction image obtained for comparison: {patient.Id}, {patient.LastName}, {pd_plan.Id}, {PDimage.Image.CreationDateTime}, {pd_beam.Id}");
                                        }
                                    }
                                }
                                number_of_analyzed_objects++;
                            }
                            catch
                            {
                                if (pd_beam.Beam.CreationDateTime > start_date)

                                {
                                    Console.ForegroundColor = ConsoleColor.DarkGray;
                                    Console.WriteLine($"Dose planes is missing for: {patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_beam.Id}");
                                }

                            }


                            #region Write all anlysis to a file located at folder \\Variancom\va_transfer\PD_Logs

                            string userLogPath;
                            StringBuilder userLogCsvContent = new StringBuilder();
                            if (ResultAverageDD != 0 && ResultGammaLessThanOne != 0 && ResultmaxdoseDiffRelative != 0 && ResultMaxGamma != 0)
                            {
                                if (Directory.Exists(@"\\Variancom\va_transfer\PD_Logs"))
                                {
                                    userLogPath = @"\\Variancom\va_transfer\PD_Logs\PD-Mining\" + System.DateTime.Now.ToString("yyyy-MM-dd") + "_PD-Mining.csv";
                                }
                                else
                                {
                                    userLogPath = Path.GetTempFileName() + "_" + System.DateTime.Now.ToString("yyyy-MM-dd") + "_PD - Mining.csv";
                                }


                                // add headers if the file doesn't exist

                                if (!File.Exists(userLogPath))
                                {
                                    List<string> dataHeaderList = new List<string>();
                                    dataHeaderList.Add("ID");
                                    dataHeaderList.Add("FirstName");
                                    dataHeaderList.Add("LastName");
                                    dataHeaderList.Add("Course");
                                    dataHeaderList.Add("PlanID");
                                    dataHeaderList.Add("MaxDoseDiff");
                                    dataHeaderList.Add("AverageDoseDiff");
                                    dataHeaderList.Add("MaxGamma");
                                    dataHeaderList.Add("GammaLessThenOne");

                                    string concatDataHeader = string.Join(",", dataHeaderList.ToArray());

                                    userLogCsvContent.AppendLine(concatDataHeader);
                                }


                                List<object> userStatsList = new List<object>();


                                userStatsList.Add(patient.Id);
                                userStatsList.Add(patient.FirstName.Replace(",", ""));
                                userStatsList.Add(patient.LastName.Replace(",", ""));
                                userStatsList.Add(pd_plan.PlanSetup.Course.Id.Replace(",", ""));
                                userStatsList.Add(pd_plan.Id.Replace(",", ""));
                                userStatsList.Add(ResultmaxdoseDiffRelative.ToString().Replace(",", "."));
                                userStatsList.Add(ResultAverageDD.ToString().Replace(",", "."));
                                userStatsList.Add(ResultMaxGamma.ToString().Replace(",", "."));
                                userStatsList.Add(ResultGammaLessThanOne.ToString().Replace(",", "."));

                                string concatUserStats = string.Join(",", userStatsList.ToArray());

                                userLogCsvContent.AppendLine(concatUserStats);

                                File.AppendAllText(userLogPath, userLogCsvContent.ToString(), Encoding.Unicode);

                                #endregion
                            }




                        }
                        
                    }
                    
                    number_of_patients_analyzed++;
                }



            }
            else
            {
                foreach (var patient_summary in application.PatientSummaries.Reverse())
                {
                    application.ClosePatient();
                    Patient patient = application.OpenPatient(patient_summary);



                    if (patient != null)
                    {


                        foreach (PDPlanSetup pd_plan in patient.PDPlanSetups.OrderByDescending(z => z.HistoryDateTime))
                        {


                            foreach (PDBeam pd_beam in pd_plan.Beams.OrderByDescending(z => z.HistoryDateTime))


                            //Where(x=>x.Beam.CreationDateTime>start_date)
                            {
                                iterator++;
                                number_of_analyzed_objects++;
                                if (iterator > 10000000)

                                {
                                    Console.WriteLine("\n\n\nThis program was made an attempt to analyze more than 1 million objects.\nExecution is interrupted for safety purposes!\nTo exit the application, please, press ENTER...");
                                    Console.ReadLine();
                                    goto Finished;
                                }

                                //define metrics to show to a user
                                double ResultmaxdoseDiffRelative = 0;
                                double ResultAverageDD = 0;
                                double ResultMaxGamma = 0;
                                double ResultGammaLessThanOne = 0;
                                try
                                {


                                    List<EvaluationTestDesc> evaluationTestDescs = new List<EvaluationTestDesc>();
                                    EvaluationTestDesc Desc_MaxDose = new EvaluationTestDesc(EvaluationTestKind.MaxDoseDifferenceRelative, double.NaN, 1, false);
                                    EvaluationTestDesc Desc_AverageDD = new EvaluationTestDesc(EvaluationTestKind.AverageDoseDifferenceRelative, double.NaN, 3, false);
                                    EvaluationTestDesc Desc_MaxGamma = new EvaluationTestDesc(EvaluationTestKind.MaxGamma, double.NaN, 3, false);
                                    EvaluationTestDesc Desc_GammaLessThanOne = new EvaluationTestDesc(EvaluationTestKind.GammaAreaLessThanOne, double.NaN, 100, false);
                                    evaluationTestDescs.Add(Desc_MaxDose);
                                    evaluationTestDescs.Add(Desc_AverageDD);
                                    evaluationTestDescs.Add(Desc_MaxGamma);
                                    evaluationTestDescs.Add(Desc_GammaLessThanOne);
                                    //evaluationTestDescs.Add(Desk_GammaLessThanOne);

                                    PDTemplate templatePD = new PDTemplate(false, false, false, false, AnalysisMode.Relative, NormalizationMethod.MinimizeDifference, true, 0.2, ROIType.CIAO, 5, GammaParameterDoseDifference, GammaParameterDistanceToAgreement, false, evaluationTestDescs);


                                    PortalDoseImage image_to_analyze = pd_beam.PortalDoseImages.Last();
                                    DoseImage reference_image = pd_beam.PortalDoseImages.First();
                                    DoseImage baseline_image = pd_beam.ConstancyCheckBaselineImage;

                                    if (baseline_image != null)
                                    {
                                        reference_image = baseline_image;
                                    }

                                    PDAnalysis analysis = new PDAnalysis();



                                    //Lets use try/catch to avoid malfunctions
                                    foreach (PortalDoseImage PDimage in pd_beam.PortalDoseImages.OrderByDescending(z => z.HistoryDateTime))
                                    {
                                        if (PDimage.Image.CreationDateTime > start_date)
                                        {
                                            try
                                            {

                                                analysis = PDimage.CreateTransientAnalysis(templatePD, reference_image);

                                                EvaluationTest TestmaxdoseDiffRelative = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxDoseDifferenceRelative);
                                                ResultmaxdoseDiffRelative = Math.Round(TestmaxdoseDiffRelative.TestValue * 100, 2);


                                                EvaluationTest TestAverageDD = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.AverageDoseDifferenceRelative);
                                                ResultAverageDD = Math.Round(TestAverageDD.TestValue * 100, 2);

                                                EvaluationTest TestMaxGamma = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.MaxGamma);
                                                ResultMaxGamma = Math.Round(TestMaxGamma.TestValue, 2);

                                                EvaluationTest TestGammaLessThanOne = analysis.EvaluationTests.FirstOrDefault(x => x.EvaluationTestKind == EvaluationTestKind.GammaAreaLessThanOne);
                                                ResultGammaLessThanOne = Math.Round(TestGammaLessThanOne.TestValue * 100, 2);

                                                //Fulfill lists with the data to be exported to excel
                                                ID_List.Add(patient.Id);
                                                FN_List.Add(patient.FirstName);
                                                LN_List.Add(patient.LastName);
                                                Course_List.Add(pd_plan.PlanSetup.Course.Id);
                                                PlanID_List.Add(pd_plan.Id);
                                                BeamID_List.Add(pd_beam.Id);
                                                TreatedDate_List.Add(Convert.ToDateTime(PDimage.Image.CreationDateTime));
                                                MDD_List.Add(ResultmaxdoseDiffRelative);
                                                ADD_List.Add(ResultAverageDD);
                                                MG_List.Add(ResultMaxGamma);
                                                GLTO_List.Add(ResultGammaLessThanOne);

                                                if (ResultGammaLessThanOne < 90 && ResultMaxGamma > 2)
                                                {
                                                    Console.BackgroundColor = default;
                                                    Console.BackgroundColor = ConsoleColor.DarkRed;
                                                    Console.ForegroundColor = ConsoleColor.Gray;
                                                    Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_beam.Id}, {pd_plan.PlanSetup.Course.Id}, {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                    Console.BackgroundColor = default;
                                                }
                                                else if ((pd_plan.Id.Contains("H") && pd_plan.Id.Contains("N")) && (ResultAverageDD > 1 && ResultMaxGamma > 1.5))
                                                {
                                                    Console.BackgroundColor = default;
                                                    Console.BackgroundColor = ConsoleColor.DarkMagenta;
                                                    Console.ForegroundColor = ConsoleColor.Gray;
                                                    Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                    Console.BackgroundColor = default;
                                                }
                                                else if (ResultmaxdoseDiffRelative < 1 && ResultAverageDD < 1 && ResultMaxGamma < 1)
                                                {
                                                    Console.BackgroundColor = default;
                                                    Console.ForegroundColor = ConsoleColor.DarkGray;
                                                    Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, Probably, it is a portal QA field or reference dose image");
                                                    Console.BackgroundColor = default;
                                                }
                                                else if (ResultmaxdoseDiffRelative>80)
                                                {
                                                    Console.BackgroundColor = default;
                                                    Console.BackgroundColor = ConsoleColor.DarkCyan;
                                                    Console.ForegroundColor = ConsoleColor.Gray;
                                                    Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_beam.Id}, {pd_plan.PlanSetup.Course.Id}, {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                    Console.BackgroundColor = default;
                                                }
                                                else if ((pd_plan.Id.Contains("Brain") || (pd_plan.Id.Contains("Brn"))) && (ResultmaxdoseDiffRelative > 10 || ResultAverageDD > 1 || ResultMaxGamma > 2))
                                                {
                                                    Console.BackgroundColor = default;
                                                    Console.BackgroundColor = ConsoleColor.DarkMagenta;
                                                    Console.ForegroundColor = ConsoleColor.Gray;
                                                    Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_plan.PlanSetup.Course.Id}, {pd_beam.Id},  {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                    Console.BackgroundColor = default;
                                                }
                                                else
                                                {
                                                    Console.BackgroundColor = default;
                                                    Console.ForegroundColor = ConsoleColor.Gray;
                                                    Console.WriteLine($"{patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_beam.Id}, {pd_plan.PlanSetup.Course.Id}, {PDimage.Image.CreationDateTime}, MaxDoseDiff: {ResultmaxdoseDiffRelative},  AverageDoseDiff: {ResultAverageDD}, MaxGamma: {ResultMaxGamma}, GammaLessThanOne: {ResultGammaLessThanOne}");
                                                    Console.BackgroundColor = default;
                                                }

                                                counter_is_the_patient_is_counted = 1;
                                            }
                                            catch
                                            {
                                                Console.ForegroundColor = ConsoleColor.DarkGray;
                                                Console.WriteLine($"No first-fraction image obtained for comparison: {patient.Id}, {patient.LastName}, {pd_plan.Id}, {PDimage.Image.CreationDateTime}, {pd_beam.Id}");
                                            }
                                        }
                                    }
                                        
                                }
                                catch
                                {
                                    if (pd_beam.Beam.CreationDateTime > start_date)

                                    {
                                        Console.ForegroundColor = ConsoleColor.DarkGray;
                                        Console.WriteLine($"Dose planes is missing for: {patient.Id}, {patient.LastName}, {pd_plan.Id}, {pd_beam.Id}");
                                    }

                                }


                                #region Write all anlysis to a file located at folder \\Variancom\va_transfer\PD_Logs

                                string userLogPath;
                                StringBuilder userLogCsvContent = new StringBuilder();
                                if (ResultAverageDD != 0 && ResultGammaLessThanOne != 0 && ResultmaxdoseDiffRelative != 0 && ResultMaxGamma != 0)
                                {
                                    if (Directory.Exists(@"\\Variancom\va_transfer\PD_Logs"))
                                    {
                                        userLogPath = @"\\Variancom\va_transfer\PD_Logs\PD-Mining\" + patientID + System.DateTime.Now.ToString("yyyy-MM-dd") + "_PD-Mining.csv";
                                    }
                                    else
                                    {
                                        userLogPath = Path.GetTempFileName() + patientID + System.DateTime.Now.ToString("yyyy-MM-dd") + "_PD - Mining.csv";
                                    }


                                    // add headers if the file doesn't exist

                                    if (!File.Exists(userLogPath))
                                    {
                                        List<string> dataHeaderList = new List<string>();
                                        dataHeaderList.Add("ID");
                                        dataHeaderList.Add("FirstName");
                                        dataHeaderList.Add("LastName");
                                        dataHeaderList.Add("Course");
                                        dataHeaderList.Add("PlanID");
                                        dataHeaderList.Add("MaxDoseDiff");
                                        dataHeaderList.Add("AverageDoseDiff");
                                        dataHeaderList.Add("MaxGamma");
                                        dataHeaderList.Add("GammaLessThenOne");

                                        string concatDataHeader = string.Join(",", dataHeaderList.ToArray());

                                        userLogCsvContent.AppendLine(concatDataHeader);
                                    }


                                    List<object> userStatsList = new List<object>();


                                    userStatsList.Add(patient.Id);
                                    userStatsList.Add(patient.FirstName.Replace(",", ""));
                                    userStatsList.Add(patient.LastName.Replace(",", ""));
                                    userStatsList.Add(pd_plan.PlanSetup.Course.Id.Replace(",", ""));
                                    userStatsList.Add(pd_plan.Id.Replace(",", ""));
                                    userStatsList.Add(ResultmaxdoseDiffRelative.ToString().Replace(",", "."));
                                    userStatsList.Add(ResultAverageDD.ToString().Replace(",", "."));
                                    userStatsList.Add(ResultMaxGamma.ToString().Replace(",", "."));
                                    userStatsList.Add(ResultGammaLessThanOne.ToString().Replace(",", "."));

                                    string concatUserStats = string.Join(",", userStatsList.ToArray());

                                    userLogCsvContent.AppendLine(concatUserStats);

                                    File.AppendAllText(userLogPath, userLogCsvContent.ToString(), Encoding.Unicode);

                                    #endregion
                                }




                            }
                        }

                        if (counter_is_the_patient_is_counted == 1)
                        {
                            number_of_patients_analyzed++;
                            counter_is_the_patient_is_counted = 0;
                        }
                    }

                    //number_of_patients_analyzed++;
                    number_of_patients_database++;

                }
            }

            Finished:
            application.ClosePatient();

            #region This module will write result to an Excel file and will save it to the directory ExcelUserLogPath defined above


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            excelApp.Visible = false;
            excelApp.UserControl = false;
            workbook = excelApp.Workbooks.Add();
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            string[] ID_Array = ID_List.ToArray();
            string[] FN_Array = FN_List.ToArray();
            string[] LN_Array = LN_List.ToArray();
            string[] Course_Array = Course_List.ToArray();
            string[] PlanID_Array = PlanID_List.ToArray();
            string[] BeamID_Array = BeamID_List.ToArray();
            DateTime[] TreatedDate_Array = TreatedDate_List.ToArray();
            double[] MDD_Array = MDD_List.ToArray();
            double[] ADD_Array = ADD_List.ToArray();
            double[] MG_Array = MG_List.ToArray();
            double[] GLTO_Array = GLTO_List.ToArray();

            //Fulfill Excel worksheet with the data obtained (z is the index of a row)
            int number_of_beams_analyzed = ID_Array.Length;
            for (int z = 2; z <= number_of_beams_analyzed; z++)
            {
                worksheet.Cells[z, 1] = ID_Array[z-1];
                worksheet.Cells[z, 2] = LN_Array[z-1];
                worksheet.Cells[z, 3] = FN_Array[z-1];
                worksheet.Cells[z, 4] = Course_Array[z-1];
                worksheet.Cells[z, 5] = PlanID_Array[z-1];
                worksheet.Cells[z, 6] = BeamID_Array[z - 1];
                worksheet.Cells[z, 7] = TreatedDate_Array[z - 1];
                worksheet.Cells[z, 8] = MDD_Array[z-1];
                worksheet.Cells[z, 9] = ADD_Array[z-1];
                worksheet.Cells[z, 10] = MG_Array[z-1];
                worksheet.Cells[z, 11] = GLTO_Array[z-1];
            }
            worksheet.Cells[1, 1] = "PatientID";
            worksheet.Cells[1, 2] = "LastName";
            worksheet.Cells[1, 3] = "FirstName";
            worksheet.Cells[1, 4] = "CourseID";
            worksheet.Cells[1, 5] = "PlanID";
            worksheet.Cells[1, 6] = "BeamID";
            worksheet.Cells[1, 7] = "Treated at:";
            worksheet.Cells[1, 8] = "MaxDoseDiff(%)";
            worksheet.Cells[1, 9] = "AverageDoseDiff(%)";
            worksheet.Cells[1, 10] = "MaxGamma";
            worksheet.Cells[1, 11] = "AreaGammaLessThanOne";
            worksheet.Cells[1, 12] = "DTA";
            worksheet.Cells[2, 12] = Convert.ToString(GammaParameterDoseDifference*100);
            worksheet.Cells[1, 13] = "GammaDoseDiff(%)";
            worksheet.Cells[2, 13] = Convert.ToString(GammaParameterDoseDifference*100);
            
            worksheet.Cells.EntireColumn.AutoFit();



            object FileNameRoute;
            if (specific_patient > 1) 
            {
                application.ClosePatient();
                Patient patient = application.OpenPatientById(patientID);

                try { FileNameRoute = @"\\Variancom\va_transfer\PD_Logs\PD-Mining\PD" + "_" + patient.LastName + "_" + patient.Id +"_"+GammaParameterDoseDifference*100+"_"+GammaParameterDistanceToAgreement+"_"+System.DateTime.Now.ToString("yyyy-MM-dd"); }
                catch { Console.WriteLine("\n\n\nMalfunction!\nPlease, check patient's ID and accessibility  of instance: 'Variancom/va/transfer/PD_Logs/PD-Mining/PD'\n\n\n\n\n\n\n\n");}
                FileNameRoute = @"\\Variancom\va_transfer\PD_Logs\PD-Mining\PD" + "_" + patient.LastName + "_" + patient.Id + "_"+ GammaParameterDoseDifference*100 + "_" + GammaParameterDistanceToAgreement + "_" + System.DateTime.Now.ToString("yyyy-MM-dd");
                workbook.SaveAs(FileNameRoute, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                workbook.Close(true, Type.Missing, Type.Missing);
                excelApp.Quit();
            }
            else
            {
                FileNameRoute = @"\\Variancom\va_transfer\PD_Logs\PD-Mining\PD_Report" +"_"+ GammaParameterDoseDifference*100 + "_" + GammaParameterDistanceToAgreement + "_" + System.DateTime.Now.ToString("yyyy-MM-dd");

                workbook.SaveAs(FileNameRoute, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                workbook.Close(true, Type.Missing, Type.Missing);
                excelApp.Quit();
            }
            
            #endregion
            
            
            //To show notification after the execution is over
            var execution_time = Math.Round((((DateTime.Now - stopwatch_start).TotalSeconds) / 60));
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Dear, {0}.\nThe analysis took {1} minutes\n", username, execution_time);
            Console.WriteLine("Gamma parameters are: {0}%/{1}mm", GammaParameterDoseDifference*100,GammaParameterDistanceToAgreement);
            Console.WriteLine("You can find report at (Variancom/va_transferPD_Logs/PD-Mining)\n");
            Console.WriteLine("The number of analyzed objects is {0}\n", number_of_analyzed_objects);
            Console.WriteLine("The number of patients successfuly analyzed is {0}\n", number_of_patients_analyzed);
            Console.WriteLine("The number of patients in the database is {0}\n\n\n", number_of_patients_database+1);
            Console.WriteLine("If you would like to stop execution of this program,\nplease, press 'ENTER'");
            Console.WriteLine("");
            Console.ReadLine();

        }

    }

}




