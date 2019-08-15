using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;

using System.Windows.Forms;

using System.Data.SqlClient;

using System.Data.OleDb;
using System.Text.RegularExpressions;


namespace Old_OfCommon
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try{
                string str_ofstring = textBox1.Text.Trim();
                int int_site = str_ofstring.LastIndexOf('_');
                string str_SOLDTO = str_ofstring.Substring(0, int_site);
                string str_Runmode = str_ofstring.Substring(str_ofstring.Length-1,1);              
                clsGlobalVar.SetDBConn(str_SOLDTO, str_Runmode);



                string strSQL = "SELECT @@SERVERNAME";
                string dtSHPMNT = clsGlobalVar.objSQLAccess.GetFieldValue(strSQL);
                MessageBox.Show(dtSHPMNT);
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string REFNM4 = "PO10301398_-00010";
            int A = Regex.Matches(REFNM4, @"_").Count;
        }

        private void button3_Click(object sender, System.EventArgs e)
        {

        }

        private void button4_Click(object sender, System.EventArgs e)
        {
            try
            {

                clsGlobalVar.SetDBConn("ss");

                string strSMTPServer = "SDSAP07";
                string str_Message = "";
              
                clsGlobalVar.objMailAccess.SendMail(clsGlobalVar.arrDefaultMail, strSMTPServer, "QCI Richard's Testing", str_Message, null);




            }
            catch(Exception ex)
            {


               
                throw ex;

            }
        }

        private void button5_Click(object sender, System.EventArgs e)
        {
            
            string str_Path = @"C:\Users\A7036325\Desktop\TEST111.xls";
            DataTable dtUpload = clsGlobalVar.objFileAccess.ExceltoDataSet(str_Path, true).Tables[0];

        }

        private void button6_Click(object sender, System.EventArgs e)
        {
            try
            {
          //<add name="NONEDI_V_T" connectionString="Initial Catalog=SDS_NONEDI;Data Source=sds8;Integrated Security=False;Password=SDSNONEDI;User ID=NONEDI;"providerName="System.Data.SqlClient" />  

               string str_SOLDTO="NONEDI_V";
               string   str_Runmode="T";
                clsGlobalVar.SetDBConn(str_SOLDTO, str_Runmode);



                string strSQL = "SELECT @@SERVERNAME";
                string dtSHPMNT = clsGlobalVar.objSQLAccess.GetFieldValue(strSQL);
                MessageBox.Show(dtSHPMNT);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void button7_Click(object sender, System.EventArgs e)
        {

            string checkString = textBox1.Text.ToString();
            for (int counter = 0; counter < checkString.Length; counter++)
            {
                if (2 * checkString[counter].ToString().Length == Encoding.Default.GetByteCount(checkString[counter].ToString()))
                {
                    MessageBox.Show("full width" + checkString[counter].ToString() + "[" + (counter+1) + "]");
                }

            }
        }
    }
}
