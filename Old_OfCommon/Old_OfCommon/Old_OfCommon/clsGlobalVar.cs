using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.Data;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text.RegularExpressions;


using QCI.OF_Common;
namespace Old_OfCommon
{
    public class clsGlobalVar
    {
        public static QCI.OF_Common.SqlClient objSQLAccess;
        public static QCI.OF_Common.Mail objMailAccess;
        public static QCI.OF_Common.FileAccess objFileAccess;
        public static QCI.OF_Common.FTPFactory objFTP;

        public static string strRegionNow;

        public static string strSenderID;
        public static string strReceiverID;
        public static string strGSNum;
        public static string strGSID;
        public static string strISA7;
         //P=Production, T=Test
        //public static string strClient = "218";
        public static string strClient = "QA2";

        public static string strSMTPServer = "172.20.166.182";

        public static string strDBLinkSoldTo = "QSMCONEGDS";
        public static string strRunMode = "T";

        //public static string[] arrRegions = { "CQ-Packing", "Packing-SDS", "Packing", "CQ-QMPacking" }; //QCMC,QSMC,CSMC,QCMCQM
        public static string[] arrRegions = { "Packing-SDS" }; //QSMC


        public static string[] alDefaultMail1 = { "Jinwei.Chang@quantacn.com" }; //QSMC


        public static ArrayList arrDefaultMail = new ArrayList(alDefaultMail1);


        public static void SetDBConn()
        {
            try
            {
                QCI.OF_Common.InitRemoting.SetSqlClient();
                QCI.OF_Common.InitRemoting.SetFileAccess();
                QCI.OF_Common.InitRemoting.SetMail();
                objSQLAccess = QCI.OF_Common.InitRemoting.objSqlClient;
                objSQLAccess.SetDBConnection(strDBLinkSoldTo, strRunMode);
                objSQLAccess.CommandTimeout = 3000;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        

        public static void SetDBConn(string strRegion)
        {

            

            string strRegion1 = "QSMCONEGDS";
            string strRunMode = "T";
            InitRemoting.SetMail();
            objMailAccess = InitRemoting.objMail;
            objMailAccess.SetDefaultMailTo(new ArrayList(arrDefaultMail));
            InitRemoting.SetSqlClient();
            objSQLAccess = InitRemoting.objSqlClient;
            objSQLAccess.SetDBConnection(strRegion1, strRunMode);
            objSQLAccess.CommandTimeout = 0;


            InitRemoting.SetFileAccess();
            objFileAccess = InitRemoting.objFileAccess;
        }

        public static void SetDBConn(string strRegion,string strRunMode)
        {
            try
            {
                QCI.OF_Common.InitRemoting.SetSqlClient();
                QCI.OF_Common.InitRemoting.SetFileAccess();
                QCI.OF_Common.InitRemoting.SetMail();
                objSQLAccess = QCI.OF_Common.InitRemoting.objSqlClient;
                objSQLAccess.SetDBConnection(strRegion, strRunMode);
                objSQLAccess.CommandTimeout = 3000;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string getSite(string strRegion)
        {
            string strSite="";
            switch(strRegion){
                case "CQ-Packing":
                    strSite = "QCMC";
                    strClient = "218";
                    break;
                case "Packing-SDS":
                    strSite = "QSMC";
                    strClient = "QA2";
                    break;
                case "Packing":
                    strSite = "CSMC";
                    strClient = "218";
                    break;
                case "CQ-QMPacking":
                    strSite = "CQQM";
                    strClient = "228";
                    break;
                default:
                    strSite = "";
                    break;
            }
            return strSite;
        }

        public static void SetMailConn()
        {
            QCI.OF_Common.InitRemoting.SetMail();
            objMailAccess = QCI.OF_Common.InitRemoting.objMail;
        }

        public static ArrayList GetDefaultMailList()
        {
            ArrayList arrMailList = new ArrayList();
            return arrMailList;
        }

        public static void SendMail(string strTitle, string strContent, ArrayList Attachment, string strGroup)
        {
            SetMailConn();
            SetDBConn();
            ArrayList strOut = new ArrayList();
            strOut.Add("Stone.Shi@quantacn.com");
            strOut.Add("Claire.Wang@quantacn.com");
            clsGlobalVar.objMailAccess.SendMail(strOut, clsGlobalVar.strSMTPServer, strTitle, strContent, Attachment, "HTML", "Web_Notice@quantacn.com");

        }


        public static void SendMail_QCMC(string strTitle, string strContent, ArrayList Attachment, string strGroup)
        {
            SetMailConn();
            ArrayList strOut = new ArrayList();
            strOut.Add("stone.shi@quantacn.com");
            strOut.Add("Claire.Wang@quantacn.com");

            clsGlobalVar.objMailAccess.SendMail(strOut, clsGlobalVar.strSMTPServer, strTitle, strContent, Attachment, "HTML", "Web_Notice@quantacn.com");

        }


        public static bool DTtoFile(DataTable dtUI, string strPath)
        {
            string strContent = "";
            try
            {
                StreamWriter sw = new StreamWriter(strPath, false,Encoding.UTF8);

                for (int i = 0; i < dtUI.Rows.Count; i++)
                {
                    for (int j = 0; j < dtUI.Columns.Count; j++)
                    {
                        strContent += dtUI.Rows[i][j].ToString();
                        if (j < dtUI.Columns.Count - 1)
                        {
                            strContent += "\t";
                        }
                    }
                    strContent += "\r\n";
                }

                sw.Write(strContent);
                sw.Close();
                return true;
            }
            catch (Exception Ex)
            {
                return false;
            }
        }

        public static bool DTtoFileSpecial(DataTable dtUI, string strPath)
        {
            string strContent = "";
            try
            {
                StreamWriter sw = new StreamWriter(strPath, false, Encoding.UTF8);

                strContent += "<RECORD>\t" + "\r\n";

                for (int i = 0; i < dtUI.Rows.Count; i++)
                {
                    for (int j = 0; j < dtUI.Columns.Count; j++)
                    {
                        strContent += dtUI.Rows[i][j].ToString();
                        strContent += "\t";
                    }
                    strContent += "\r\n";
                }

                sw.Write(strContent);
                sw.Close();
                return true;
            }
            catch (Exception Ex)
            {
                return false;
            }
        }
    }
}
