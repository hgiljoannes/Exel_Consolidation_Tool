using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;

namespace Exel_Consolidation_Tool.ExcelProcess
{
    public class killExcel
    {
        public void KillExcel()
        {
            try
            {
                Process[] processes = Process.GetProcessesByName("EXCEl");
                int pid = processes[0].Id;
                Process pro = Process.GetProcessById(pid);
                pro.Kill();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Killing Process " + ex.Message);
                throw;
            }
          
        }
    }
}
