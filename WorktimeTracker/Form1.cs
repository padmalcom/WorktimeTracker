using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Globalization;
using Microsoft.Win32;
using System.Reflection;
using WorktimeTracker.Properties;

namespace WorktimeTracker
{
    public partial class Form1 : Form
    {

        private static string APPLICATION_DIRECTORY = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\WorktimeTracker\\";
        private static string EXCEL_FILE = APPLICATION_DIRECTORY + "Zeiterfassung.xls";
        private static string LAST_DAY_FILE = APPLICATION_DIRECTORY + "lastDay.txt";
        private static string APPLICATION_NAME = "WorktimeTracker";
        private static int WM_QUERYENDSESSION = 0x11;
        private static bool systemShutdown = false;
        private static bool endTimeWasWritten = false;
        private static bool rpmeReminder11 = false;
        private static bool rpmeReminder14 = false;
        private static DateTime currentDay = DateTime.Today;
        private static DateTime suspendTime, suspendDay;

        public Form1()
        {
            InitializeComponent();
        }

        // Shutdown hook
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            if (m.Msg == WM_QUERYENDSESSION)
            {
                systemShutdown = true;
            }
            base.WndProc(ref m);
        }

        // React on resize to change notify icon
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                notifyIcon1.Visible = true;
                notifyIcon1.ShowBalloonTip(500);
                this.Hide();
            }
            else if (FormWindowState.Normal == this.WindowState)
            {
                notifyIcon1.Visible = false;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.
            Opacity = 0;

            base.OnLoad(e);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SystemEvents.PowerModeChanged += OnPowerChange;

            // Create the application directory if not exists
            if (!Directory.Exists(APPLICATION_DIRECTORY))
            {
                Directory.CreateDirectory(APPLICATION_DIRECTORY);
            }

            // Does the file exist?
            if (!File.Exists(EXCEL_FILE))
            {
                MessageBox.Show("Zeiterfassung-Sheet existiert nicht. Erstelle '" + EXCEL_FILE + "'.");
                File.WriteAllBytes(EXCEL_FILE, Properties.Resources.Zeiterfassung);
            }

            newDay();

        }

        private void newDay()
        {
            // On PC start -> ask to register start
            registerStart(false);


            // Has an end date been written in the last days?
            if (File.Exists(LAST_DAY_FILE))
            {
                int year, month, day, hour, minute;
                string monthStr;
                if (readEndText(out year, out month, out day, out hour, out minute, out monthStr))
                {
                    // Register end
                    if (MessageBox.Show("Es wurde für den " + day + "." + month + "." + year + " kein Arbeitsende eingetragen. Soll" +
                        "dieses auf den Runterfahrzeitpunkt des PCs (" + hour + ":" + minute + ") gesetzt werden?", "Arbeitsende", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        registerEnd(true, day, year, month, hour, minute, monthStr);
                        File.Delete(LAST_DAY_FILE);
                    }
                }
            }
        }

        private void registerStart(bool clicked)
        {
            int day = DateTime.Now.Day;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int hours = DateTime.Now.Hour;
            int minutes = DateTime.Now.Minute;
            string monthStr = DateTime.Now.ToString("MMM", CultureInfo.InvariantCulture);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = null;
            Worksheet ws = null;

            // Open file
            if (xlApp == null)
            {
                MessageBox.Show("Excel ist nicht installiert. Beende.");
                return;
            }

            if (IsFileinUse(new FileInfo(EXCEL_FILE)))
            {
                MessageBox.Show("Kann '" + EXCEL_FILE + "' nicht bearbeiten, da die Datei bereits geöffnet ist.");
                return;
            }

            if (File.Exists(EXCEL_FILE))
            {
                wb = xlApp.Workbooks.Open(EXCEL_FILE);
            }
            else
            {
                wb = xlApp.Workbooks.Add();
                wb.SaveAs(EXCEL_FILE);
            }

            // Open or create sheet
            bool exists = false;
            Worksheet wsOriginal = null;
            foreach (Worksheet currentWs in wb.Worksheets)
            {
                if (currentWs.Name == monthStr + "_" + year)
                {
                    exists = true;
                    ws = currentWs;
                    break;
                }
                if (currentWs.Name == "Original")
                {
                    wsOriginal = currentWs;
                }
            }
            if (!exists)
            {
                if (wsOriginal == null)
                {
                    MessageBox.Show("Es wurde keine Sheet mit Namen 'Original' gefunden. Um die Anwendung eine neue Excel-Datei erstellen zu lassen,"+
                        " löschen Sie bitte '" + EXCEL_FILE + "'. Nach dem nächsten Start dieser Anwendung wird ein intaktes Sheet angelegt.");
                    return;
                }
                wsOriginal.Copy(Type.Missing, wb.Sheets[wb.Sheets.Count]);
                ws = wb.Sheets[wb.Sheets.Count];
                ws.Name = monthStr + "_" + year;
                ws.Cells[23, 3].Value = month;
                ws.Cells[23, 4].Value = year;
            }

            int startCell = 23 + day;
            int endCell = 23 + day;

            //return;

            string currentDaysStart = ws.Cells[startCell, 3].Text;
            string currentDaysEnd = ws.Cells[endCell, 4].Text;
                        
            string wsDay = ws.Cells[23 + day, 1].Text;

            if (wsDay != "Sa" || wsDay != "So")
            {
                if (currentDaysStart == "07:00" || clicked)
                {
                    if (!clicked)
                    {
                        if (MessageBox.Show("Möchten Sie den Beginn Ihrer Arbeitszeit eintragen?", "Arbeitsbeginn", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            ws.Cells[23 + day, 3].Value = hours + ":" + minutes;
                        }
                    }
                    else
                    {
                        ws.Cells[23 + day, 3].Value = hours + ":" + minutes;
                    }
                }
            }
            else
            {
                if (currentDaysStart == "00:00" || clicked)
                {
                    if (!clicked)
                    {
                        if (MessageBox.Show("Möchten Sie den Beginn Ihrer Arbeitszeit eintragen?", "Arbeitsbeginn", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            ws.Cells[23 + day, 3].Value = hours + ":" + minutes;
                        }
                    }
                    else
                    {
                        ws.Cells[23 + day, 3].Value = hours + ":" + minutes;
                    }
                }
            }
            wb.Save();
            wb.Close();
            xlApp.Quit();
        }

        private void registerEnd(bool clicked, int day, int year, int month, int hours, int minutes, string monthStr)
        {
            /*int day = DateTime.Now.Day;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int hours = DateTime.Now.Hour;
            int minutes = DateTime.Now.Minute;
            string monthStr = DateTime.Now.ToString("MMM", CultureInfo.InvariantCulture);*/

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = null;
            Worksheet ws = null;

            // Open file
            if (xlApp == null)
            {
                MessageBox.Show("Excel ist nicht installiert. Beende.");
                return;
            }

            if (IsFileinUse(new FileInfo(EXCEL_FILE)))
            {
                MessageBox.Show("Kann '" + EXCEL_FILE + "' nicht bearbeiten, da die Datei bereits geöffnet ist.");
                return;
            }

            if (File.Exists(EXCEL_FILE))
            {
                wb = xlApp.Workbooks.Open(EXCEL_FILE);
            }
            else
            {
                wb = xlApp.Workbooks.Add();
                wb.SaveAs(EXCEL_FILE);
            }

            // Open or create sheet
            bool exists = false;
            Worksheet wsOriginal = null;
            foreach (Worksheet currentWs in wb.Worksheets)
            {
                if (currentWs.Name == monthStr + "_" + year)
                {
                    exists = true;
                    ws = currentWs;
                    break;
                }
                if (currentWs.Name == "Original")
                {
                    wsOriginal = currentWs;
                }
            }
            if (!exists)
            {
                if (wsOriginal == null)
                {
                    MessageBox.Show("Es wurde keine Sheet mit Namen 'Original' gefunden. Um die Anwendung eine neue Excel-Datei erstellen zu lassen," +
                        " löschen Sie bitte '" + EXCEL_FILE + "'. Nach dem nächsten Start dieser Anwendung wird ein intaktes Sheet angelegt.");
                    return;
                }
                wsOriginal.Copy(Type.Missing, wb.Sheets[wb.Sheets.Count]);
                ws = wb.Sheets[wb.Sheets.Count];
                ws.Name = monthStr + "_" + year;
                ws.Cells[23, 3].Value = month;
                ws.Cells[23, 4].Value = year;
            }

            string currentDaysEnd = ws.Cells[23 + day, 4].Text;

            string wsDay = ws.Cells[23 + day, 1].Text;

            if (wsDay != "Sa" || wsDay != "So")
            {
                if (!clicked)
                {
                    if (MessageBox.Show("Möchten Sie das Ende Ihres Arbeitstages eintragen?", "Arbeitsende", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ws.Cells[23 + day, 4].Value = hours + ":" + minutes;
                    }
                }
                else
                {
                    ws.Cells[23 + day, 4].Value = hours + ":" + minutes;
                }
            }
            else
            {
                if (!clicked)
                {
                    if (MessageBox.Show("Möchten Sie das Ende Ihres Arbeitstages eintragen?", "Arbeitsende", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ws.Cells[23 + day, 4].Value = hours + ":" + minutes;
                    }
                }
                else
                {
                    ws.Cells[23 + day, 4].Value = hours + ":" + minutes;
                }
            }
            wb.Save();
            wb.Close();
            xlApp.Quit();
            endTimeWasWritten = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (systemShutdown)
            {
                if (!endTimeWasWritten)
                {
                    writeEndText();
                }

            }
        }

        private void beginWorkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            registerStart(true);
        }

        private void endWorkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int day = DateTime.Now.Day;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int hours = DateTime.Now.Hour;
            int minutes = DateTime.Now.Minute;
            string monthStr = DateTime.Now.ToString("MMM", CultureInfo.InvariantCulture);
            registerEnd(true, day, year, month, hours, minutes, monthStr);
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (!toolStripMenuItem1.Checked)
            {
                registryKey.SetValue(APPLICATION_NAME, System.Windows.Forms.Application.ExecutablePath);
            }
            else
            {
                registryKey.DeleteValue(APPLICATION_NAME);
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (registryKey.GetValue(APPLICATION_NAME) == null)
            {
                toolStripMenuItem1.Checked = false;
            }
            else
            {
                toolStripMenuItem1.Checked = true;
            }
        }

        private bool IsFileinUse(FileInfo file)
        {

            if (!file.Exists) return false;

            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(EXCEL_FILE);
        }

        public static void writeEndText() {
            string[] lines = { DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.ToString("MMM", CultureInfo.InvariantCulture)};
            File.WriteAllLines(LAST_DAY_FILE, lines);
        }

        public static bool readEndText(out int year, out int month, out int day, out int hour, out int minute, out string monthStr)
        {
            string line;
            year = 0;
            month = 0;
            day = 0;
            hour = 0;
            minute = 0;
            monthStr = "";
            System.IO.StreamReader file = new StreamReader(LAST_DAY_FILE);
            line = file.ReadLine();
            if (line != null) { year = int.Parse(line); } else return false;

            line = file.ReadLine();
            if (line != null) { month = int.Parse(line); }  else return false;

            line = file.ReadLine();
            if (line != null) { day = int.Parse(line); }  else return false;

            line = file.ReadLine();
            if (line != null) { hour = int.Parse(line); }  else return false;

            line = file.ReadLine();
            if (line != null) { minute = int.Parse(line); } else return false;

            line = file.ReadLine();
            if (line != null) { monthStr = line; } else return false;

            file.Close();

            return true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            int dayInMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);

            // Remind user in the last 2 month's days
            if (DateTime.Now.Day >= dayInMonth - 4 && DateTime.Now.Day <= dayInMonth)
            {
                if (DateTime.Now.Hour == 11 && DateTime.Now.Minute == 0 && DateTime.Now.Second == 0 && rpmeReminder11 == false)
                {
                    rpmeReminder11 = true;
                    notifyIcon1.BalloonTipTitle = "RPME";
                    notifyIcon1.BalloonTipText = "Bitte denken Sie daran Ihre Zeitem im RPME einzutragen.";
                    notifyIcon1.ShowBalloonTip(5);
                }
                else if (DateTime.Now.Hour == 14 && DateTime.Now.Minute == 0 && DateTime.Now.Second == 0 && rpmeReminder14 == false)
                {
                    rpmeReminder14 = true;
                    notifyIcon1.BalloonTipTitle = "RPME";
                    notifyIcon1.BalloonTipText = "Bitte denken Sie daran Ihre Zeitem im RPME einzutragen.";
                    notifyIcon1.ShowBalloonTip(5);
                }
            }

            if (DateTime.Today > currentDay)
            {
                currentDay = DateTime.Today;
                newDay();
            }
        }

        private void OnPowerChange(object s, PowerModeChangedEventArgs e)
        {
            switch (e.Mode)
            {
                case PowerModes.Resume:
                    if (suspendDay == DateTime.Today.AddDays(-1))
                    {
                        if (MessageBox.Show("Der PC wurde gestern (" + suspendTime.Day + "." + suspendTime.Month + "." + suspendTime.Year + ") um " +
                            "" + suspendTime.Hour + ":" + suspendTime.Minute + " Uhr suspended. Soll diese Zeit als Arbeitsende eingetragen werden?", "Arbeitsende", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            registerEnd(true, suspendTime.Day, suspendTime.Year, suspendTime.Month, suspendTime.Hour, suspendTime.Minute, suspendTime.ToString("MMM", CultureInfo.InvariantCulture));
                    }
                    break;
                case PowerModes.Suspend:
                    suspendTime = DateTime.Now;
                    suspendDay = DateTime.Today;
                    break;
            }
        }
    }

}
