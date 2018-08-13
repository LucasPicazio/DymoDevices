using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;

namespace USBdevices
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        bool TemCardScan(string BRD)
        {
            try
            {
                if(BRD != "")
                {
                    ManagementObjectSearcher searcher =
                                        new ManagementObjectSearcher("\\\\" + BRD + "\\root\\CIMV2",
                                        "SELECT * FROM Win32_USBControllerDevice");

                    foreach (ManagementObject queryObj in searcher.Get())
                    {
                        if (queryObj["Dependent"].ToString().Contains("VID_08F0&PID_0005")) return true;

                    }
                    return false;
                }
                return false;
            }
            catch (Exception x)
            {
                
                return false;
            }
        }

     

        private void tb_KeyDown(object sender, KeyPressEventArgs e)
        {
            
            if (Char.IsControl(e.KeyChar) && e.KeyChar != '\b')
            {
                
                int count = 0;
                foreach (var line in richTextBox1.Lines)
                {
                    try
                    {
                        if(line != "")
                        {
                            dataGridView1.Rows.Add(line, TemCardScan(line), TemDymoLabel(line));
                            count++;
                        }
                        label1.Text = count.ToString() + "/" + richTextBox1.Lines.Count().ToString();
                    }
                    catch(ManagementException s)
                    {
                        count++;
                        label1.Text = count.ToString()+"/"+ richTextBox1.Lines.Count().ToString();
                    }
                }

            }
        }

        private object TemDymoLabel(string BRD)
        {
            try
            {
                if (BRD != "")
                {
                    ManagementObjectSearcher searcher =
                                        new ManagementObjectSearcher("\\\\" + BRD + "\\root\\CIMV2",
                                        "SELECT * FROM Win32_USBControllerDevice");

                    foreach (ManagementObject queryObj in searcher.Get())
                    {
                        if (queryObj["Dependent"].ToString().Contains("DYMO")) return true;
                    }
                    return false;
                }
                return false;
            }
            catch (Exception x)
            {

                return false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            richTextBox1.KeyPress += new KeyPressEventHandler(tb_KeyDown);
        }
    }
}
