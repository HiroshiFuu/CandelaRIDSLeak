/*
 * Candela RIDS Leak Tester
 * Seethoo Wai Leong
 * 
 * Hardware:
 *  AtEQ F520
 *  Adlink PCI-1733 32 Channel Digital Inputs
 *  Adlink PCI-1734 32 Channel Digital Outputs
 * 
 * Version tested with all on engineering page working : 201410101132
 * Note that serial is 7 databits.
 * 
 * Added Test Suite and Test Parameters.
 * Added super-functions
 * 
*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using Automation.BDaq;
using System.Threading.Tasks;

namespace CandelaRIDSLeak
{
    public partial class Form1 : Form
    {
        //long cylinderTimeAllowed;
        int AteQTimeAllowed;
        bool goodToContinue;
        bool safetyCurtainCrossed;
        bool leftCylinderPos;
        bool last_safetyCurtainCrossed_1, last_safetyCurtainCrossed_2;
        bool last_leftIsUp, last_leftIsDown;
        bool last_safetyCurtainCrossed, last_leftCylinderPos;
         
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                serialPort1.Open();
            }

            catch
            {
                MessageBox.Show("Port connected is NOT COM1: Please check.");


                logging("failed startup", "none");
                Application.Exit();
                return;
            }
            
            // make these user settings
            //cylinderTimeAllowed = 5000;
            AteQTimeAllowed = 25000;
            resetTest();

            leftDown(false);
            rightIn(false);
            rightDown(false);

            resetLabels();

            groupBox1.Refresh();

            instantDiCtrl1.SelectedDevice = new DeviceInformation("PCI-1733,BID#0");
            instantDoCtrl1.SelectedDevice = new DeviceInformation("PCI-1734,BID#0");
            
            //The default device of project is demo device, users can choose other devices according to their needs. 
            if (!instantDoCtrl1.Initialized)
            {
                MessageBox.Show("No device be selected or device open failed!", "StaticDO");
                this.Close();
                return;
            }

            if (!instantDiCtrl1.Initialized)
            {
                MessageBox.Show("No device be selected or device open failed!", "StaticDI");
                this.Close();
                return;
            }

            this.Text = "( " + instantDoCtrl1.SelectedDevice.Description + " + " + instantDiCtrl1.SelectedDevice.Description + ")";

            // AteQ start pulse pull down on start up
            switchOff(0, 1);

            logAction("initial");
            timer2.Start();
        }

        void resetLabels()
        {
            resultLeakAll.Text = "";
            resultLeakY.Text = "";
            resultLeakM.Text = "";
            resultLeakC.Text = "";
            resultLeakK.Text = "";
            resultContY.Text = "";
            resultContM.Text = "";
            resultContC.Text = "";
            resultContK.Text = "";
            resultLeakAll.BackColor = Color.Gray;
            resultLeakY.BackColor = Color.Gray;
            resultLeakM.BackColor = Color.Gray;
            resultLeakC.BackColor = Color.Gray;
            resultLeakK.BackColor = Color.Gray;
            resultContY.BackColor = Color.Gray;
            resultContM.BackColor = Color.Gray;
            resultContC.BackColor = Color.Gray;
            resultContK.BackColor = Color.Gray;
        }

        private void resetTest()
        {
            goodToContinue = true;
            safetyCurtainCrossed = last_safetyCurtainCrossed = false;
            last_safetyCurtainCrossed_1 = last_safetyCurtainCrossed_2 = false;
            leftCylinderPos = last_leftCylinderPos = true;
            last_leftIsUp = last_leftIsDown = true;
            timer1.Start();
            logEvent("resetTest", "goodToContinue", goodToContinue.ToString());
        }

        #region Logging Func
        private void logging(string remarks, string para)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@".\log.txt", true))
            {
                file.WriteLine(remarks + " @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                if (remarks.CompareTo("testing") == 0)
                    file.WriteLine("testPassed " + para);
                file.WriteLine("goodToContinue " + goodToContinue);
                file.WriteLine("  Y    M    C    K");
                file.WriteLine(chkLeakTestY.Checked + " " + chkLeakTestM.Checked + " " + chkLeakTestC.Checked + " " + chkLeakTestK.Checked);
                file.WriteLine();
            }
        }

        private void logEvent(string ev, string para, string paraV)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@".\log.txt", true))
            {
                file.WriteLine(ev + " @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                file.WriteLine(para + " " + paraV);
                file.WriteLine();
            }
        }

        private void logAction(string act)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@".\log.txt", true))
            {
                file.WriteLine(act + " @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                file.WriteLine();
            }
        }

        private void logTimer()
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@".\logTimer.txt", true))
            {
                file.WriteLine("Timer.Enabled is " + timer1.Enabled + " @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                file.WriteLine();
            }
        }

        private void logLeftCylinder(string ev, string para, string paraV)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@".\leftCylinder.txt", true))
            {
                file.WriteLine(ev + " @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                file.WriteLine(para + " " + paraV);
                file.WriteLine();
            }
        }

        private void logSafetyCurtain(string ev, string para, string paraV)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@".\safetyCurtain.txt", true))
            {
                file.WriteLine(ev + " @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                file.WriteLine(para + " " + paraV);
                file.WriteLine();
            }
        }
        #endregion 

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPageIndex == 2)
            {
                serialPort1.Close();
                Application.Exit();
                return;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (leftButtonPressed() && rightButtonPressed())
            {
                logAction("start test from button");
                runworker();
            }

            bool safetyCurtainCrossed1 = safetyCurtainCrossed_1();
            bool safetyCurtainCrossed2 = safetyCurtainCrossed_2();

            if (estopIsPressed()) label28.BackColor = Color.Green; else label28.BackColor = Color.Gray;
            if (leftButtonPressed()) label21.BackColor = Color.Green; else label21.BackColor = Color.Gray;
            if (rightButtonPressed()) label22.BackColor = Color.Green; else label22.BackColor = Color.Gray;
            if (safetyCurtainCrossed1) label26.BackColor = Color.Green; else label26.BackColor = Color.Gray;
            if (safetyCurtainCrossed2) label27.BackColor = Color.Green; else label27.BackColor = Color.Gray;

            //goodToContinue = !(safetyCurtainCrossed_1() || safetyCurtainCrossed_2() || estopIsPressed());
            safetyCurtainCrossed = (safetyCurtainCrossed1 || safetyCurtainCrossed2);
            if (last_safetyCurtainCrossed != safetyCurtainCrossed || last_safetyCurtainCrossed_1 != safetyCurtainCrossed1 || last_safetyCurtainCrossed_2 != safetyCurtainCrossed2)
            {
                last_safetyCurtainCrossed = safetyCurtainCrossed;
                last_safetyCurtainCrossed_1 = safetyCurtainCrossed1;
                last_safetyCurtainCrossed_2 = safetyCurtainCrossed2;
                logEvent("safetyCurtainCrossed", "safetyCurtainCrossed 1/2", safetyCurtainCrossed1 + "/" + safetyCurtainCrossed2);
            }
            bool leftIs_Up = leftIsUp();
            bool leftIs_Down = leftIsDown();
            leftCylinderPos = !(leftIs_Up || leftIs_Down);
            if (last_leftCylinderPos != leftCylinderPos || last_leftIsUp != leftIs_Up || last_leftIsDown != leftIs_Down)
            {
                last_leftCylinderPos = leftCylinderPos;
                last_leftIsUp = leftIs_Up;
                last_leftIsDown = leftIs_Down;
                logEvent("leftCylinderPos", "leftIs Up/Down", leftIs_Up + "/" + leftIs_Down);
            }
            if (safetyCurtainCrossed && leftCylinderPos)
                emergencyStop();

            // Cylinders positions
            if (leftIs_Up) chkLeftCylinder.BackColor = Color.Green;
            if (leftIs_Down) chkLeftCylinder.BackColor = Color.Red; 
            if (rightIsUp()) chkRightCylinder.BackColor = Color.Green; 
            if (rightIsDown()) chkRightCylinder.BackColor = Color.Red; 
            if (rightIsRetracted()) chkNeedleCylinder.BackColor = Color.Green; 
            if (rightIsInserted()) chkNeedleCylinder.BackColor = Color.Red; 

            if (AteQ_Passed()) label23.BackColor = Color.Green; 
            if (AteQ_Failed()) label24.BackColor = Color.Green; 
            if (AteQ_AlarmRaised()) label24.BackColor = Color.Green;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            logTimer();
            bool safetyCurtainCrossed1 = safetyCurtainCrossed_1();
            bool safetyCurtainCrossed2 = safetyCurtainCrossed_2();
            logSafetyCurtain("safetyCurtainCrossed", "safetyCurtainCrossed 1/2", safetyCurtainCrossed1 + "/" + safetyCurtainCrossed2);
            bool leftIs_Up = leftIsUp();
            bool leftIs_Down = leftIsDown();
            logLeftCylinder("leftCylinderPos", "leftIs Up/Down", leftIs_Up + "/" + leftIs_Down);
        }

        #region Fundamental Lowlevel Functions

        byte bitSet(int position) // zero based
        {
            int mask = 1;
            mask <<= position;
            return (byte) mask;
        }

        byte bitClear(int position)
        {
            int mask = 1;
            mask <<= position;
            return (byte) ~mask;
        }

        ErrorCode switchOn(int portNum, int bitNum)
        {
            byte statusByte = 0;
            byte mask = 0;
            ErrorCode err = instantDoCtrl1.Read(portNum, out statusByte);
            if (err == ErrorCode.Success)
            {
                mask = bitSet(bitNum);
                statusByte |= mask;
                err = instantDoCtrl1.Write(portNum, statusByte);
            }
            return err;
        }

        ErrorCode switchOff(int portNum, int bitNum)
        {
            byte statusByte = 0;
            byte mask = 0;
            ErrorCode err = instantDoCtrl1.Read(portNum, out statusByte);
            if (err == ErrorCode.Success)
            {
                mask = bitClear(bitNum);
                statusByte &= mask;
                err = instantDoCtrl1.Write(portNum, statusByte);
            }
            return err;
        }

        Boolean isOn(int portNum, int bitNum)
        {
            byte statusByte = 0;
            bool result = false;
            ErrorCode err = instantDiCtrl1.Read(portNum, out statusByte);
            if (err == ErrorCode.Success)
            {
                result = isBitOn(statusByte, bitNum);
            }

            return result;
        }

        Boolean isBitOn(byte statusByte, int bitNum)
        {
            byte mask = bitSet(bitNum);
            statusByte &= mask;
            statusByte >>= bitNum;

            return (statusByte == 1);
        }

        #endregion

        #region Digital Inputs and Outputs made simple with names

        void activateCylinder(int port0utput, int bitOutput, int portSensorHome, int bitSensorHome, int portSensorExtended, int bitSensorExtended, bool extend)
        {

            if (chkEnableCylinders.Checked == true)
            {
                if (extend) switchOn(port0utput, bitOutput);
                else switchOff(port0utput, bitOutput);
                return;
                if (chkEnableSmartCylinders.Checked == true)
                // wait until sensor is triggered to indicate required cylinder is reached.
                {
                    //Stopwatch timer;
                    //timer = Stopwatch.StartNew();
                    if (extend)
                    {
                        while (!isOn(portSensorExtended, bitSensorExtended))
                        {
                            //timer.Stop();
                            //if (timer.ElapsedMilliseconds > cylinderTimeAllowed) MessageBox.Show("Timed Out");
                            //throw new TimeoutException();
                            //timer.Start();
                            Application.DoEvents();
                        }
                    }
                    else
                    {
                        while (!isOn(portSensorHome, bitSensorHome))
                        {
                            //timer.Stop();
                            //if (timer.Elapsed.Seconds > cylinderTimeAllowed) MessageBox.Show("Timed Out");
                            //throw new TimeoutException();
                            //timer.Start();
                            Application.DoEvents();
                        }
                    }
                }
            }
            
        }

        void leftDown(bool trueIfDownFalseIfUp)
        {
            activateCylinder(0, 3, 0, 3, 0, 4, trueIfDownFalseIfUp);
        }

        void rightDown(bool trueIfDownFalseIfUp)
        {
            activateCylinder(0, 4, 0, 5, 0, 6, trueIfDownFalseIfUp);
        }

        void rightIn(bool trueIfInFalseIfOut)
        {
            // Needle can only be pushed in when the right cylinder is clamped down.
            // Needle can only retracted when cylinder is still down.
            if ((chkEnableSmartCylinders.Checked && rightIsDown()) || !chkEnableSmartCylinders.Checked)
            {
                activateCylinder(0, 5, 0, 7, 1, 0, trueIfInFalseIfOut);
            }
        }

        //void leftDown(bool trueIfDownFalseIfUp)
        //{
        //    if (chkEnableCylinders.Checked == true)
        //    {
        //        if (trueIfDownFalseIfUp) switchOn(0, 3);
        //        else switchOff(0, 3);
        //    }
        //}
        
        //void rightDown(bool trueIfDownFalseIfUp)
        //{
        //    if (chkEnableCylinders.Checked == true)
        //    {
        //        if (trueIfDownFalseIfUp) switchOn(0, 4);
        //        else switchOff(0, 4);
        //    }

        //}

        //void rightIn(bool trueIfInFalseIfOut)
        //{
        //    if (chkEnableCylinders.Checked == true)
        //    {
        //        if (trueIfInFalseIfOut) switchOn(0, 5);
        //        else switchOff(0, 5);
        //    }
        //}

        void startAteQ()  // its a pulse
        {
            switchOn(0, 1);
            Thread.Sleep(100);  // maybe shorter
            switchOff(0, 1);
        }

        void selectContinuityProgram()
        {
            switchOn(0, 2);
            switchOff(1, 1);
        }

        void selectLeakProgram()
        {
            switchOff(0, 2);
            switchOn(1, 1);
        }

        void redLightOn(bool trueIfOnFalseIfOff)
        {
            if (trueIfOnFalseIfOff) switchOn(0, 6);
            else switchOff(0, 6);
        }

        void yellowLightOn(bool trueIfOnFalseIfOff)
        {
            if (trueIfOnFalseIfOff) switchOn(0, 7);
            else switchOff(0, 7);
        }

        void greenLightOn(bool trueIfOnFalseIfOff)
        {
            if (trueIfOnFalseIfOff) switchOn(1, 0);
            else switchOff(1, 0);
        }

        void valveOpen(int valveNumber, bool trueIfValveIsToBeOpened)
        {
            int port = 2;
            if (valveNumber < 6)
            {
                port = 1;
                valveNumber += 2;
            }
            else
            {
                valveNumber -= 6;
            }

            if (trueIfValveIsToBeOpened) switchOn(port, valveNumber);
            else switchOff(port, valveNumber);
        }

        void valveAll_InputOpen()
        {
            byte mask = 0x78;   // 0111 1000
            byte statusByte = 0;
            ErrorCode err = instantDoCtrl1.Read(1, out statusByte);

            if (err == ErrorCode.Success)
            {
                statusByte |= mask;
                err = instantDoCtrl1.Write(1, statusByte);
            }  
        }

        void valveAll_InputClose()
        {
            byte mask = 0x87;   // 1000 0111
            byte statusByte = 0;
            ErrorCode err = instantDoCtrl1.Read(1, out statusByte);

            if (err == ErrorCode.Success)
            {
                statusByte &= mask;
                err = instantDoCtrl1.Write(1, statusByte);
            }  
        }

        void valveAll_OutputOpen()
        {
            switchOn(1, 7);
            
            byte mask = 0x07;   // 0000 0111
            byte statusByte = 0;
            ErrorCode err = instantDoCtrl1.Read(2, out statusByte);

            if (err == ErrorCode.Success)
            {
                statusByte |= mask;
                err = instantDoCtrl1.Write(2, statusByte);
            }
        }

        void valveAll_OutputClose()
        {
            switchOff(1, 7);

            byte mask = 0xf8;   // 1111 1000
            byte statusByte = 0;
            ErrorCode err = instantDoCtrl1.Read(2, out statusByte);

            if (err == ErrorCode.Success)
            {
                statusByte &= mask;
                err = instantDoCtrl1.Write(2, statusByte);
            }
        }

        bool leftIsUp()
        {
            return isOn(0, 3);
        }
        
        bool leftIsDown()
        {
            return isOn(0, 4);
        }

        bool rightIsUp()
        {
            return isOn(0, 5);
        }

        bool rightIsDown()
        {
            return isOn(0, 6);
        }

        bool rightIsRetracted()
        {
            return isOn(0, 7);
        }

        bool rightIsInserted()
        {
            return isOn(1, 0);
        }

        bool leftButtonPressed()
        {
            return isOn(0, 1);
        }

        bool rightButtonPressed()
        {
            return isOn(0, 2);
        }

        bool estopIsPressed()
        {
            return !isOn(0, 0);
        }

        bool safetyCurtainCrossed_1()
        {
            return isOn(1, 1);
        }

        bool safetyCurtainCrossed_2()
        {
            return isOn(1, 2);
        }

        bool AteQ_Passed()
        {
            return isOn(1, 5);
        }

        bool AteQ_Failed()
        {
            return isOn(1, 6);
        }

        bool AteQ_AlarmRaised()
        {
            return isOn(1, 7);
        }

        private bool Part_Presence()
        {
            return isOn(2, 1);
        }

        #endregion

        #region Engineering Mode Buttons

        #region EventHandler
        private void btnRightIn_Click(object sender, EventArgs e)
        {
            rightIn(true);
        }

        private void btnRightOut_Click(object sender, EventArgs e)
        {
            rightIn(false);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            byte portStatus = 0;
            ErrorCode err = instantDiCtrl1.Read(0, out portStatus);
            button24.Text = portStatus.ToString("X2");
        }

        private void button25_Click(object sender, EventArgs e)
        {
            byte portStatus = 0;
            ErrorCode err = instantDiCtrl1.Read(1, out portStatus);
            button25.Text = portStatus.ToString("X2");
        }

        private void button26_Click(object sender, EventArgs e)
        {
            byte portStatus = 0;
            ErrorCode err = instantDiCtrl1.Read(2, out portStatus);
            button26.Text = portStatus.ToString("X2");
        }

        private void button27_Click(object sender, EventArgs e)
        {
            byte portStatus = 0;
            ErrorCode err = instantDiCtrl1.Read(3, out portStatus);
            button27.Text = portStatus.ToString("X2");
        }

        private void btnSelectLeak_Click(object sender, EventArgs e)
        {
            selectLeakProgram();
        }

        private void btnSelectContinuity_Click(object sender, EventArgs e)
        {
            selectContinuityProgram();
        }

        private void btnStartAteqTest_Click(object sender, EventArgs e)
        {
            startAteQ();
        }

        private void chkLeftCylinder_CheckedChanged(object sender, EventArgs e)
        {

            if (chkLeftCylinder.Checked) leftDown(true);
            else leftDown(false);

        }

        private void chkRightCylinder_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRightCylinder.Checked) rightDown(true);
            else rightDown(false);
        }

        private void chkNeedleCylinder_CheckedChanged(object sender, EventArgs e)
        {
            if (chkNeedleCylinder.Checked) rightIn(true);
            else rightIn(false);
        }

        private void chkRedLight_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRedLight.Checked) redLightOn(true);
            else redLightOn(false);
        }

        private void chkYellowLight_CheckedChanged(object sender, EventArgs e)
        {
            if (chkYellowLight.Checked) yellowLightOn(true);
            else yellowLightOn(false);
        }

        private void chkGreenLight_CheckedChanged(object sender, EventArgs e)
        {
            if (chkGreenLight.Checked) greenLightOn(true);
            else greenLightOn(false);
        }

        private void chkValve1_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve1.Checked) valveOpen(1, true);
            else valveOpen(1, false);            
        }

        private void chkValve2_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve2.Checked) valveOpen(2, true);
            else valveOpen(2, false);  
        }

        private void chkValve3_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve3.Checked) valveOpen(3, true);
            else valveOpen(3, false);  
        }

        private void chkValve4_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve4.Checked) valveOpen(4, true);
            else valveOpen(4, false);  
        }

        private void chkValve5_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve5.Checked) valveOpen(5, true);
            else valveOpen(5, false); 
        }

        private void chkValve6_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve6.Checked) valveOpen(6, true);
            else valveOpen(6, false);
        }

        private void chkValve7_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve7.Checked) valveOpen(7, true);
            else valveOpen(7, false); 
        }

        private void chkValve8_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValve8.Checked) valveOpen(8, true);
            else valveOpen(8, false); 
        }

        private void chkValvesAll_Input_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValvesAll_Input.Checked) valveAll_InputOpen();
            else valveAll_InputClose();
        }

        private void chkValvesAll_Output_CheckedChanged(object sender, EventArgs e)
        {
            if (chkValvesAll_Output.Checked) valveAll_OutputOpen();
            else valveAll_OutputClose();
        }


        private void chkLeakTestEach_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLeakTestEach.Checked)
            {
                chkLeakTestY.Checked = true;
                chkLeakTestM.Checked = true;
                chkLeakTestC.Checked = true;
                chkLeakTestK.Checked = true;
            }
            else
            {
                chkLeakTestY.Checked = false;
                chkLeakTestM.Checked = false;
                chkLeakTestC.Checked = false;
                chkLeakTestK.Checked = false;
            }
        }

        private void chkContinuityTest_CheckedChanged(object sender, EventArgs e)
        {
            if (chkContinuityTest.Checked)
            {
                chkContinuityTestY.Checked = true;
                chkContinuityTestM.Checked = true;
                chkContinuityTestC.Checked = true;
                chkContinuityTestK.Checked = true;

            }
            else
            {
                chkContinuityTestY.Checked = false;
                chkContinuityTestM.Checked = false;
                chkContinuityTestC.Checked = false;
                chkContinuityTestK.Checked = false;
            }
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                if (serialPort1.BytesToRead > 0)
                {
                    txtSerial.AppendText("0123456789012345678901234567890\n");
                    txtSerial.AppendText(serialPort1.ReadExisting());
                }
            }
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            saveParameters();
        }

        private void btnStartTest_Click(object sender, EventArgs e)
        {
            logAction("start test from GUI");
            runworker();
        }

        #endregion

        /// <summary>
        /// Superfunction to test a channel for leak
        /// </summary>
        /// <param name="channel">Channel 1 to 4 for YMCA. 0 if all channels are to be tested.</param>
        /// <returns></returns>
        /// 
        bool leakTest(int channel)
        {
            
            bool isGood = true;
            Label label = resultLeakAll;
            float leakLimit = Convert.ToInt16(txtLeakLimitAll.Text);
            if (channel == 0)
            {
                valveAll_InputOpen();
                label = resultLeakAll;
            }

            else
            {
                valveOpen(channel, true);
                switch (channel)
                {
                    case 1:
                        label = resultLeakY; break;
                    case 2:
                        label = resultLeakM; break;
                    case 3:
                        label = resultLeakC; break;
                    case 4:
                        label = resultLeakK; break;
                }
                leakLimit = Convert.ToSingle(txtLeakLimitEach.Text);
            }

            label.ForeColor = Color.Black;
            label.Text = "Testing...";
            label.BackColor = Color.Yellow;
            label.Refresh();

            selectLeakProgram();
            //System.Threading.Thread.Sleep(1000);  // Does timer interrupt works during sleep? FH: No, Thread.Sleep will block the program.
            waits(1000);
            startAteQ();

            String testPressure;
            String result;
            isGood = analyseAteQOutput(out testPressure, out result);
            if (isGood)
            {
                isGood = (Convert.ToSingle(result) < leakLimit);
            }

            valveAll_InputClose();


            lblTestPressure.Text = testPressure;
            label.Text = result;

            // Change color
            label.ForeColor = Color.White;
            if ( isGood )
            {
                label.BackColor = Color.Green;
            }
            else
            {
                label.BackColor = Color.Red;
            }

            label.Refresh();
            return isGood;
        }

        bool continuityTest(int channel)
        {
            Label label = resultContC;
            String testPressure;
            String result;

            bool isGood = true;
            valveAll_InputClose();
            valveAll_OutputClose();
            selectContinuityProgram();
            //System.Threading.Thread.Sleep(1000);
            waits(1000);

            switch (channel)
            {
                case 1:
                    valveOpen(1, true);
                    valveOpen(8, true);
                    label = resultContY; break;

                case 2:
                    valveOpen(2, true);
                    valveOpen(7, true);
                    label = resultContM; break;

                case 3:
                    valveOpen(3, true);
                    valveOpen(6, true);
                    label = resultContC; break;

                case 4:
                    valveOpen(4, true);
                    valveOpen(5, true);
                    label = resultContK; break;

            }

            label.BackColor = Color.Yellow;
            label.Refresh();


            startAteQ();

            isGood = !analyseAteQOutput(out testPressure, out result);

            // bad means there is a leak, which is what we want in the continuity tester, so bad is good
            if (isGood)
            {
                // isGood = (Convert.ToInt16(result) > Convert.ToInt16(txtContinuityTestThreshold.Text));
                isGood = (Convert.ToSingle(testPressure) < Convert.ToSingle(txtContinuityTestThreshold.Text));
            }

            label.Text = testPressure;

            if (isGood)
            {
                label.BackColor = Color.Green;
            }
            else
            {
                label.BackColor = Color.Red;
            }

            valveAll_InputClose();
            valveAll_OutputClose();
            label.Refresh();

            return isGood;
        }

        private void waits(int mili)
        {
            Stopwatch sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 1000)
            {
                Application.DoEvents();
            }
        }

        void standBy(long milliseconds)
        {
            Stopwatch sw = Stopwatch.StartNew();
        }

        bool analyseAteQOutput(out string testPressure, out string result)
        {
            bool returnVal = true;
            testPressure = "N.A.";
            result = "N.A.";
            try
            {
                if (serialPort1.IsOpen)
                {
                    Stopwatch timer;
                    timer = Stopwatch.StartNew();
                    while (serialPort1.BytesToRead == 0) // we have a problem is we don't have anything!
                    {
                        timer.Stop();
                        if (timer.ElapsedMilliseconds > AteQTimeAllowed) MessageBox.Show("analyseAteQOutput Timed Out");
                        //throw new TimeoutException();
                        timer.Start();
                    }

                    string response = serialPort1.ReadLine();
                    txtSerial.AppendText(response + "\n");
                    testPressure = response.Substring(6, 5); // need to check this
                    
                    //First significant response comes on the character 5 and 6. C# is zero based.
                    if (response.IndexOf("AL")>-1)  //
                    {
                        returnVal = false;
                    }
                    else
                    {
                        result = response.Substring(22, 3);
                    }
                }
            }
            catch
            {
                MessageBox.Show("No response from AteQ detected.");
                returnVal = false;
            }
            return returnVal;
        }

        void startTest()
        {
            // All test begins with the left cylinder being press down first.
            timer1.Enabled = true;
            logEvent("Test started", "timer1.Enabled", timer1.Enabled.ToString());
            bool testPassed = false;
            greenLightOn(true);
            leftDown(true);
            rightDown(true);
            while (!leftIsDown())
            {
                Application.DoEvents();
            }
            timer1.Enabled = false;
            logEvent("preparation done", "timer1.Enabled", timer1.Enabled.ToString());

            resetLabels();
            lblFinalResult.Text = "Testing...";
            lblFinalResult.ForeColor = Color.Black;
            lblFinalResult.BackColor = Color.Gray;
            lblFinalResult.Refresh();

            groupBox1.Refresh();

            logging("start leakTest(0)", testPassed.ToString());

            if (chkLeakTestAll.Checked) testPassed = leakTest(0);
    
            testPassed = true;  //all channel leak test no longer a condition for testPassed

            logging("leakTest(0) done", testPassed.ToString());

            if (!testPassed)
            {
                if (chkLeakTestY.Checked && goodToContinue) testPassed &= leakTest(1);
                if (chkLeakTestM.Checked && goodToContinue) testPassed &= leakTest(2);
                if (chkLeakTestC.Checked && goodToContinue) testPassed &= leakTest(3);
                if (chkLeakTestK.Checked && goodToContinue) testPassed &= leakTest(4);
                logEvent("chkLeakTest Done", "testPassed", testPassed.ToString());
            }

            if (chkContinuityTestY.Checked || chkContinuityTestM.Checked || chkContinuityTestC.Checked || chkContinuityTestK.Checked)
            {
                rightIn(true);
                if (chkContinuityTestY.Checked && goodToContinue) testPassed &= continuityTest(1);
                if (chkContinuityTestM.Checked && goodToContinue) testPassed &= continuityTest(2);
                if (chkContinuityTestC.Checked && goodToContinue) testPassed &= continuityTest(3);
                if (chkContinuityTestK.Checked && goodToContinue) testPassed &= continuityTest(4);
                logEvent("chkContinuityTest Done", "testPassed", testPassed.ToString());
            }
            logEvent("all tests done", "timer1.Enabled", timer1.Enabled.ToString());

            rightIn(false);
            rightDown(false);
            leftDown(false);
            lblFinalResult.ForeColor = Color.White;

            if (testPassed)
            {
                lblFinalResult.Text = "PASSED";
                lblFinalResult.BackColor = Color.Green;

            }
            else
            {
                lblFinalResult.Text = "FAILED";
                lblFinalResult.BackColor = Color.Red;
            }
            lblFinalResult.Refresh();
            timer1.Enabled = true;
            logEvent("testing done", "timer1.Enabled", timer1.Enabled.ToString());

            // open file and write out the test results
            writeResults();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            btnStartTest.Enabled = false;
            resetTest();
            last_safetyCurtainCrossed = !last_safetyCurtainCrossed;
            last_leftCylinderPos = !last_leftCylinderPos;
            startTest();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStartTest.Enabled = true;
        }

        private void runworker()
        {
            if (!Part_Presence())
            {
                MessageBox.Show("No part detected!");
                return;
            }
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
                waits(150);
                timer1.Start();
            }
            if (!backgroundWorker1.IsBusy)
                backgroundWorker1.RunWorkerAsync();
        }

        void writeResults()
        {
            // filename will be today's date including year
            string today = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");
            using (System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\testdata\\Burst_Data\\Result.tst",true))
            {
                file.Write("F9A30-60081,");
                file.Write(txtPartSerialNumber.Text);
                file.Write(",");
                file.Write("CDSRID");
                file.Write(",");
                file.Write(today);
                file.Write(",");
                file.Write(DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00"));
                file.Write(",");
                file.Write("TST201505220");
                file.Write(",");
                file.Write("CGF_NULL");
                file.Write(",");
                file.Write("18");
                file.Write(",");
                file.Write("RIDS");
                file.Write(",");
                if (lblFinalResult.Text == "PASSED") file.Write("P"); else file.Write("F");
                file.Write(",");
                file.Write(resultLeakAll.Text);
                file.Write(",");
                file.Write(resultContY.Text);
                file.Write(",");
                file.Write(resultContM.Text);
                file.Write(",");
                file.Write(resultContC.Text);
                file.Write(",");
                file.Write(resultContK.Text);
                file.Write(",");
                file.Write(resultLeakY.Text);
                file.Write(",");
                file.Write(resultLeakM.Text);
                file.Write(",");
                file.Write(resultLeakC.Text);
                file.Write(",");
                file.Write(resultLeakK.Text);

                file.WriteLine(",,,,,,,,,,,,,,,,,,,,,,,,,");
            }
        }

        void saveParameters()
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter("./CandelaRIDS.ini", true))
            {
                file.WriteLine(chkLeakTestAll.Checked);
                file.WriteLine(chkLeakTestY.Checked);
                file.WriteLine(chkLeakTestM.Checked);
                file.WriteLine(chkLeakTestC.Checked);
                file.WriteLine(chkLeakTestK.Checked);
                file.WriteLine(chkContinuityTestY.Checked);
                file.WriteLine(chkContinuityTestM.Checked);
                file.WriteLine(chkContinuityTestC.Checked);
                file.WriteLine(chkContinuityTestK.Checked);
                file.WriteLine(txtLeakLimitAll.Text);
                file.WriteLine(txtLeakLimitEach.Text);
                file.WriteLine(txtContinuityTestThreshold.Text);
            }
        }

        void emergencyStop()
        {
            goodToContinue = false;
            redLightOn(true);
            greenLightOn(false);
            leftDown(false);
            rightDown(false);

            if (backgroundWorker1.IsBusy)
                backgroundWorker1.CancelAsync();
            //System.Threading.Thread.Sleep(2000); // delay for 2 secs
            waits(2000);
            redLightOn(false);
            safetyCurtainCrossed = false;
            last_safetyCurtainCrossed = false;
            timer1.Enabled = true;
            logEvent("emergencyStop", "goodToContinue", goodToContinue.ToString());
        }

    }

}
