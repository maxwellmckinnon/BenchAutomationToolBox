using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AudioPrecision.API;
using System.IO;
using System.Runtime.InteropServices;

namespace AP_Freq_APX_FFT
{
    public partial class Form1 : Form
    {
        APx500 APx;
        AP.GlobalClass APObj;
        List<double> freqToSweep; // Hz
        string savePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //Default location
        string APConfigFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //Default location
        string dataFolder = "\\xTalkData";
        string startRunTimeString = "";

        public Form1()
        {
            InitializeComponent();
        }

        public void SetupAPSW()
        {
            //Test control of APX and AP2700 from same program -- if doable, write the rest of the program.
            
            //APX control - below code works
            textBox_Status.Text = "Latching to APx Software.";
            APx = new APx500();
            APx.Visible = true;
            //APx.SignalMonitorsEnabled = true;
            //APx.OperatingMode = APxOperatingMode.BenchMode;
            //string APxfilenameLoad = APConfigFilesPath + "\\FFT.approjx";
            //APx.OpenProject(APxfilenameLoad);

            //AP Control
            textBox_Status.Text = "Latching to AP2700 Software.";
            APObj = new AP.GlobalClass();
            APObj.Application().Visible = true;
            //string AP2700filenameLoad = APConfigFilesPath + "\\Ripple Generation.at27";
            //APObj.File().OpenTest(AP2700filenameLoad);

            //Check with user if AP2700 is working.
            //MessageBox.Show("Sometimes the AP2700 will disconnect when both APx and AP2700 softwares are loaded. Disconnect the AP2700 USB, reconnect, then click restore hardware in AP2700. Do this now before proceeding", "Crosstalk measurement via FFT");
            textBox_Status.Text = "Sometimes the AP2700 will disconnect when both APx and AP2700 softwares are loaded. Disconnect the AP2700 USB, reconnect, then click restore hardware in AP2700. Do this now before proceeding.";
        }

        private void button_SetupAPSW_Click(object sender, EventArgs e)
        {
            SetupAPSW();
        }

        private void button_SetupDefault_Click(object sender, EventArgs e)
        {
            //Setup the x axis of the crosstalk test, the frequencies to run FFTs at.
            textBox_Status.Text = "Setup the x axis of the crosstalk test, the frequencies to run FFTs at.";
            //freqToSweep = new List<double>() { 20, 30, 50, 80, 100, 217, 300, 500, 800, 1000, 2000, 3000, 5000, 8000, 10000, 15000, 20000 };
            freqToSweep = new List<double>();
            int steps = Convert.ToInt16(textBox_steps.Text);
            double maxFreq = 20000;
            double minFreq = 20;
            //Calculate logarithmic-even steps
            for (int i = 0; i < steps; i++)
            {
                double exp = Math.Pow(maxFreq/minFreq, -1 * i / (steps - 1.0));
                freqToSweep.Add(maxFreq * exp);
            }

            string freqToSweep_str = "";
            foreach (double freq in freqToSweep)
            {
                freqToSweep_str += freq.ToString() + ", ";
            }
            freqToSweep_str = freqToSweep_str.Remove(freqToSweep_str.Length - 2);
            textBox_SweepFreqs.Text = freqToSweep_str;
           
        }

        private void parse_FreqToSweep()
        {
            //Update the frequencies to sweep

            string freq_str = textBox_SweepFreqs.Text;
            freqToSweep = new List<double>(){};
            foreach (string piece in freq_str.Split(','))
            {
                freqToSweep.Add(Convert.ToDouble(piece));
            }
        }

        private void button_RunCrosstalkAutomation_Click(object sender, EventArgs e)
        {
            runCrossTalk_vs_Frequency_AP2700Ripple();

        }

        private void runCrossTalk_vs_Frequency_AP2700Ripple()
        {
            //Signal must be actively generated from the AP2700.
            //APx must be ready in FFT mode to capture the FFT

            //This part steps the frequency on the AP2700, taking an FFT of each step using the APx. Typically 1M FFT w/ 8 avg is used.

            NativeMethods.PreventSleep(); // Prevent computer from sleeping and disconnecting from APx

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            startRunTimeString = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss");

            textBox_Status.Text = "Running automation...\r\n";
            textBox_Status.Text += "Start Time: " + startRunTimeString + Environment.NewLine;
            double estimatedTimeMinutes = freqToSweep.Count * 2.25 * 60 / 30; // 30 x 8 1.2M FFT takes about 2.25 hours
            textBox_Status.Text += "Estimated Finish Time (8 averages, 1.2M length): " + DateTime.Now.AddMinutes(estimatedTimeMinutes).ToString("yyyy-MM-dd_hh-mm-ss") + Environment.NewLine;

            List<Tuple<double, double>> xTalk_freq_dBV = new List<Tuple<double, double>>();

            //Parse and update frequencies to sweep
            parse_FreqToSweep();
            
            //Begin xTalk test
            APx.BenchMode.Measurements.Fft.Append = false;
            foreach (int freq in freqToSweep)
            {
                //Change AP2700 frequency
                APObj.Gen().set_Freq("Hz", freq);
                textBox_Status.Text += freq.ToString() + "Hz" + Environment.NewLine;
                try
                {
                    //Run APx FFT
                    APx.BenchMode.Measurements.Fft.Start();
                    while(APx.BenchMode.Measurements.Fft.IsStarted) //! Wait for FFT to finish
                    {
                        System.Threading.Thread.Sleep(100); // Sleep 100ms
                    }

                    //Grab FFT data
                    double[] xValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch1);
                    double[] yValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch1);

                    double[] xValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch2);
                    double[] yValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch2);

                    //Grab desired point from data
                    double xTalkVal = findCrosstalkValue(freq, xValues, yValues, xValues_input, yValues_input);

                    //Append data to xTalk data
                    xTalk_freq_dBV.Add(new Tuple<double,double>(freq, xTalkVal));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //Export FFT data
                System.IO.Directory.CreateDirectory(savePath + dataFolder);
                string filename = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + "Freq" + freq.ToString("D5") + ".csv";
                APx.BenchMode.Measurements.Fft.ExportData(filename);

                //Export FFT Image
                string filenameJPG = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + "Freq" + freq.ToString() + ".jpg";
                APx.BenchMode.Measurements.Fft.Graphs[0].Save(filenameJPG, GraphImageType.JPG);
                //APx.BenchMode.Measurements.Fft.Graphs["Level"].Save(filenameJPG, GraphImageType.JPG); //not tested yet - test next time running code
            }

            //Export APx file - bad idea file is huge 1.5GB!
            //string filename_APx = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + ".approjx";
            //APx.SaveProject(filename_APx);

            //Export xTalk data
            exportXtalkDatatoFile(xTalk_freq_dBV);
            
            //Report test time
            sw.Stop();
            textBox_Status.Text += "Automation finished. Automation Time: " + sw.Elapsed.ToString() + Environment.NewLine + "Absolute time Ended: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + Environment.NewLine;

            NativeMethods.AllowSleep();
        }

        private void exportXtalkDatatoFile(List<Tuple<double,double>> xTalk_freq_dBV)
        {
            //string pathDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //string filePath = pathDocuments + "\\xTalkAutomation_" + "xTalkData" + startRunTimeString + ".csv";
            System.IO.Directory.CreateDirectory(savePath + dataFolder);
            string filePath = savePath + dataFolder + "\\xTalkAutomation_" + "xTalkData" + startRunTimeString + ".csv";

            if (!File.Exists(filePath))
            {
                File.Create(filePath).Close();
            }
            string delimter = ",";
            
            int length = xTalk_freq_dBV.Count;

            using (System.IO.TextWriter writer = File.CreateText(filePath))
            {
                for (int index = 0; index < length; index++)
                {
                    writer.WriteLine(string.Join(delimter, xTalk_freq_dBV[index].Item1, xTalk_freq_dBV[index].Item2));
                }
            }
        }

        private double findCrosstalkValue(double freq_Hz, double[] xValues_raw, double[] yValues_raw, double[] xValues_rawinput, double[] yValues_rawinput)
        {
            //Returns the crosstalk value for the given frequency in the data array. 
            //It is sometimes the case where the bin does not align with the frequency correctly.
                //1. Find the index closest to the freq_Hz in xValues.
                //2. Return the largest yValue around this index (add +/- 2 indices)

            double xTalkValue = 0;

            List<double> xValues = xValues_raw.ToList();
            List<double> yValues = yValues_raw.ToList();
            List<double> xValues_input = xValues_rawinput.ToList();
            List<double> yValues_input = yValues_rawinput.ToList();
            
            //1. Find index closest to freq_Hz
            int index = indexOfFreqHz(freq_Hz, xValues);
            int index_input = indexOfFreqHz(freq_Hz, xValues_input);

            //2. Return largest yValue
            double dBV = LargestAroundIndex(index, yValues);
            double dBV_input = LargestAroundIndex(index, yValues_input);

            xTalkValue = dBV - dBV_input;
            return xTalkValue;
        }

        private int indexOfFreqHz(double freq_Hz, List<double> xValues)
        {
            int index = 0;
            double oldAbsDelta = 9999999;
            double newAbsDelta = 9999999;

            //O(n) search. Probably fine
            for (int i = 0; i < xValues.Count; i++ )
            {
                newAbsDelta = Math.Abs( freq_Hz - xValues[i]);
                if (newAbsDelta < oldAbsDelta)
                {
                    oldAbsDelta = newAbsDelta;
                    index = i;
                }
            }
            return index;
        }

        private double LargestAroundIndex(int index, List<double> yValues)
        {
            //Return the largest yValue around this index (add +/- 2 indices)
            double dBV = 0;
            double newValue = -999999;
            double oldLargest = -999999;
            int searchSpread = 2;

            for (int i = index - searchSpread; i <= index + searchSpread; i++ )
            {
                try //Could possibly try to fetch an invalid index
                {
                    newValue = yValues[i];
                    if (newValue > oldLargest)
                    {
                        oldLargest = newValue;
                    }
                }
                catch
                {

                }
            }
            dBV = oldLargest;
            return dBV;
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help helpForm = new Help();
            helpForm.ShowDialog();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void setDataDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void customToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string dummyFileName = "Save data here";
            myDocumentsToolStripMenuItem.Checked = false;

            SaveFileDialog sf = new SaveFileDialog();
            // Feed the dummy name to the save dialog
            sf.FileName = dummyFileName;

            if (sf.ShowDialog() == DialogResult.OK)
            {
                savePath = Path.GetDirectoryName(sf.FileName);
            }
        }

        private void myDocumentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            savePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //Default location
            customToolStripMenuItem.Checked = false;
        }

        private void setAPConfigDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void customToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            myDocumentsToolStripMenuItem1.Checked = false;
            string dummyFileName = "Load Configuration files from here";
            myDocumentsToolStripMenuItem.Checked = false;

            SaveFileDialog sf = new SaveFileDialog();
            // Feed the dummy name to the save dialog
            sf.FileName = dummyFileName;

            if (sf.ShowDialog() == DialogResult.OK)
            {
                APConfigFilesPath = Path.GetDirectoryName(sf.FileName);
            }
        }

        private void myDocumentsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            customToolStripMenuItem1.Checked = true;
            APConfigFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //Default location
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            myDocumentsToolStripMenuItem_Click(sender, e);
            myDocumentsToolStripMenuItem.Checked = true;
            customToolStripMenuItem.Checked = false;
            myDocumentsToolStripMenuItem1_Click(sender, e);
            myDocumentsToolStripMenuItem1.Checked = true;
            customToolStripMenuItem1.Checked = false;
        }

        private void buttonDev_Click(object sender, EventArgs e)
        {
            //Put quick stuff in here. For instance, PSRR vs Voltage sweep.
            
            RunCrosstalk_vs_AVDDVoltage_APx();
            //RunCrosstalk_vs_Frequency_APx();
            
        }

        private void RunCrosstalk_vs_Frequency_APx()
        {
            //Signal generated from APx
            //APx must be ready in FFT mode to capture the FFT

            //This part steps the frequency on the AP2700, taking an FFT of each step using the APx. Typically 1M FFT w/ 8 avg is used.

            NativeMethods.PreventSleep(); // Prevent computer from sleeping and disconnecting from APx

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            startRunTimeString = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss");

            textBox_Status.Text = "Running automation...\r\n";
            textBox_Status.Text += "Start Time: " + startRunTimeString + Environment.NewLine;
            double estimatedTimeMinutes = freqToSweep.Count * 2.25 * 60 / 30; // 30 x 8 1.2M FFT takes about 2.25 hours
            textBox_Status.Text += "Estimated Finish Time (8 averages, 1.2M length): " + DateTime.Now.AddMinutes(estimatedTimeMinutes).ToString("yyyy-MM-dd_hh-mm-ss") + Environment.NewLine;

            List<Tuple<double, double>> xTalk_freq_dBV = new List<Tuple<double, double>>();

            //Parse and update frequencies to sweep
            parse_FreqToSweep();
            
            //Begin xTalk test
            APx.BenchMode.Measurements.Fft.Append = false;
            foreach (int freq in freqToSweep)
            {
                //Change APx
                APObj.Gen().set_Freq("Hz", freq);
                textBox_Status.Text += freq.ToString() + "Hz" + Environment.NewLine;
                try
                {
                    //Run APx FFT
                    APx.BenchMode.Measurements.Fft.Start();
                    while(APx.BenchMode.Measurements.Fft.IsStarted) //! Wait for FFT to finish
                    {
                        System.Threading.Thread.Sleep(100); // Sleep 100ms
                    }

                    APx.BenchMode.Measurements.Fft.Append = true;

                    //Grab FFT data
                    double[] xValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch1);
                    double[] yValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch1);

                    double[] xValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch2);
                    double[] yValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch2);

                    //Grab desired point from data
                    double xTalkVal = findCrosstalkValue(freq, xValues, yValues, xValues_input, yValues_input);

                    //Append data to xTalk data
                    xTalk_freq_dBV.Add(new Tuple<double,double>(freq, xTalkVal));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //Export FFT data
                System.IO.Directory.CreateDirectory(savePath + dataFolder);
                string filename = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + "Freq" + freq.ToString() + ".csv"; //Currently doesn't work correctly. Modify to output correct data. Currently outputs the same data over and over from the first sweep.

                APx.BenchMode.Measurements.Fft.ExportData(filename);
            }

            //Export APx file - bad idea file is huge 1.5GB!
            //string filename_APx = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + ".approjx";
            //APx.SaveProject(filename_APx);

            //Export xTalk data
            exportXtalkDatatoFile(xTalk_freq_dBV);
            
            //Report test time
            sw.Stop();
            textBox_Status.Text += "Automation finished. Automation Time: " + sw.Elapsed.ToString() + Environment.NewLine + "Absolute time Ended: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + Environment.NewLine;

            NativeMethods.AllowSleep();
        }

        private void RunCrosstalk_vs_AVDDVoltage_APx()
        {
            //Signal generated from APx
            //APx must be ready in FFT mode to capture the FFT

            //This part steps the frequency on the AP2700, taking an FFT of each step using the APx. Typically 1M FFT w/ 8 avg is used.

            NativeMethods.PreventSleep(); // Prevent computer from sleeping and disconnecting from APx

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            startRunTimeString = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss");

            textBox_Status.Text = "Running automation...\r\n";
            textBox_Status.Text += "Start Time: " + startRunTimeString + Environment.NewLine;
            
            List<double> voltageToSweep = new List<double>() {1.7, 1.75, 1.8, 1.85, 1.9, 1.95};
            PowerSupply_E3631A.E3631A PSU = new PowerSupply_E3631A.E3631A();
            double oneFreq = 1000; //Sweep voltage at 1kHz

            List<Tuple<double, double>> xTalk_voltage_dBV = new List<Tuple<double, double>>();

            //Parse and update frequencies to sweep
            //parse_FreqToSweep();
            
            //Begin xTalk test
            APx.BenchMode.Measurements.Fft.Append = false;
            foreach (double v in voltageToSweep)
            {
                //Change voltage
                PSU.setVoltage(v, "P6V");
                textBox_Status.Text += v.ToString() + "V" + Environment.NewLine;
                try
                {
                    //Run APx FFT
                    APx.BenchMode.Measurements.Fft.Start();
                    while(APx.BenchMode.Measurements.Fft.IsStarted) //! Wait for FFT to finish
                    {
                        System.Threading.Thread.Sleep(100); // Sleep 100ms
                    }

                    //APx.BenchMode.Measurements.Fft.Append = true;

                    //Grab FFT data
                    double[] xValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch1);
                    double[] yValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch1);

                    double[] xValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch2);
                    double[] yValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch2);

                    //Grab desired point from data
                    double xTalkVal = findCrosstalkValue(oneFreq, xValues, yValues, xValues_input, yValues_input);

                    //Append data to xTalk data
                    xTalk_voltage_dBV.Add(new Tuple<double,double>(v, xTalkVal));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //Export FFT data
                System.IO.Directory.CreateDirectory(savePath + dataFolder);
                string filename = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + "Voltage" + v.ToString("D5") + ".csv"; //Currently doesn't work correctly. Modify to output correct data. Currently outputs the same data over and over from the first sweep.
                APx.BenchMode.Measurements.Fft.ExportData(filename);

                //Export FFT Image
                string filenameJPG = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + "Voltage" + v.ToString("D5") + ".jpg";
                APx.BenchMode.Measurements.Fft.Graphs[0].Save(filenameJPG, GraphImageType.JPG);
                //APx.BenchMode.Measurements.Fft.Graphs["Level"].Save(filenameJPG, GraphImageType.JPG); //not tested yet - test next time running code
            }

            //Export APx file - bad idea file is huge 1.5GB!
            //string filename_APx = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + ".approjx";
            //APx.SaveProject(filename_APx);

            //Export xTalk data
            exportXtalkDatatoFile(xTalk_voltage_dBV);
            
            //Report test time
            sw.Stop();
            textBox_Status.Text += "Automation finished. Automation Time: " + sw.Elapsed.ToString() + Environment.NewLine + "Absolute time Ended: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + Environment.NewLine;

            NativeMethods.AllowSleep();
        }

        private void RunCrosstalk_vs_VBATVoltage_APx()
        {
            //Not finished

            //Signal generated from APx
            //APx must be ready in FFT mode to capture the FFT

            //This part steps the frequency on the AP2700, taking an FFT of each step using the APx. Typically 1M FFT w/ 8 avg is used.

            NativeMethods.PreventSleep(); // Prevent computer from sleeping and disconnecting from APx

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            startRunTimeString = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss");

            textBox_Status.Text = "Running automation...\r\n";
            textBox_Status.Text += "Start Time: " + startRunTimeString + Environment.NewLine;

            List<double> voltageToSweep = new List<double>() { 1.7, 1.75, 1.8, 1.85, 1.9, 1.95 };
            PowerSupply_E3631A.E3631A PSU = new PowerSupply_E3631A.E3631A();
            double oneFreq = 1000; //Sweep voltage at 1kHz

            List<Tuple<double, double>> xTalk_voltage_dBV = new List<Tuple<double, double>>();

            //Parse and update frequencies to sweep
            //parse_FreqToSweep();

            //Begin xTalk test
            APx.BenchMode.Measurements.Fft.Append = false;
            foreach (double v in voltageToSweep)
            {
                //Change AP2700 frequency
                PSU.setVoltage(v, "P6V");
                textBox_Status.Text += v.ToString() + "V" + Environment.NewLine;
                try
                {
                    //Run APx FFT
                    APx.BenchMode.Measurements.Fft.Start();
                    while (APx.BenchMode.Measurements.Fft.IsStarted) //! Wait for FFT to finish
                    {
                        System.Threading.Thread.Sleep(100); // Sleep 100ms
                    }

                    //APx.BenchMode.Measurements.Fft.Append = true;

                    //Grab FFT data
                    double[] xValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch1);
                    double[] yValues = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch1);

                    double[] xValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetXValues(InputChannelIndex.Ch2);
                    double[] yValues_input = APx.BenchMode.Measurements.Fft.FFTSpectrum.GetYValues(InputChannelIndex.Ch2);

                    //Grab desired point from data
                    double xTalkVal = findCrosstalkValue(oneFreq, xValues, yValues, xValues_input, yValues_input);

                    //Append data to xTalk data
                    xTalk_voltage_dBV.Add(new Tuple<double, double>(v, xTalkVal));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //Export FFT data
                System.IO.Directory.CreateDirectory(savePath + dataFolder);
                string filename = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + "Voltage" + v.ToString() + ".csv"; //Currently doesn't work correctly. Modify to output correct data. Currently outputs the same data over and over from the first sweep.

                APx.BenchMode.Measurements.Fft.ExportData(filename);
            }

            //Export APx file - bad idea file is huge 1.5GB!
            //string filename_APx = savePath + dataFolder + "\\xTalkAutomation_" + "rawFFTData" + startRunTimeString + ".approjx";
            //APx.SaveProject(filename_APx);

            //Export xTalk data
            exportXtalkDatatoFile(xTalk_voltage_dBV);

            //Report test time
            sw.Stop();
            textBox_Status.Text += "Automation finished. Automation Time: " + sw.Elapsed.ToString() + Environment.NewLine + "Absolute time Ended: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + Environment.NewLine;

            NativeMethods.AllowSleep();
        }

    }

    internal static class NativeMethods
    {
        //Mainly for keeping the computer from sleeping and disconnecting the APx.

        public static void PreventSleep()
        {
            SetThreadExecutionState(ExecutionState.EsContinuous | ExecutionState.EsSystemRequired);
        }

        public static void AllowSleep()
        {
            SetThreadExecutionState(ExecutionState.EsContinuous);
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);

        [FlagsAttribute]
        private enum ExecutionState : uint
        {
            EsAwaymodeRequired = 0x00000040,
            EsContinuous = 0x80000000,
            EsDisplayRequired = 0x00000002,
            EsSystemRequired = 0x00000001
        }
    }
}
