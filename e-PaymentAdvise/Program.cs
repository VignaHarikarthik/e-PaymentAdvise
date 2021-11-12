using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Mail;

namespace e_PaymentAdvise
{
    class Program
    {
        #region ***** SQL Connection*****
        static readonly SqlConnection SGConnection = new SqlConnection("Server=192.168.1.21;Database=SYSPEX_LIVE;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection JBConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Technologies (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection JKConnection = new SqlConnection("Server=192.168.1.21;Database=PT SYSPEX KEMASINDO;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection SBConnection = new SqlConnection("Server=192.168.1.21;Database=PT SYSPEX MULTITECH;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection KLConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Mechatronic (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection PGConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Industries (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static SqlConnection SAPCon12 = new SqlConnection("Server=192.168.1.21;Database=AndriodAppDB;Uid=Sa;Pwd=Password1111;");
        static readonly MySqlConnection PortalCon = new MySqlConnection("Server=192.168.1.228;Database=syspex_portal;Uid=root;Pwd=dogood;");

        static string SQLQuery;
        #endregion
        static void Main(string[] args)
        {
            //Go Live = 25/08/2020
            EPAYMENT("65ST");
            ////Go Live = 31/05/2021
            System.Threading.Thread.Sleep(5000);
            EPAYMENT("07ST");
        }

        private static void EPAYMENT(string CompanyCode)
        {
            bool flag;

            DataSet ds = GetOutNos(CompanyCode);
            if (ds.Tables[0].Rows.Count > 0)

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {

                    flag = SendInvociePDF(ds.Tables[0].Rows[i]["docnum"].ToString(),
                                          ds.Tables[0].Rows[i]["docentry"].ToString(),
                                          ds.Tables[0].Rows[i]["E_Mail"].ToString(),//Supplier email address
                                          CompanyCode, ds.Tables[0].Rows[i]["cardname"].ToString());

                    DataTable dt = CheckDuplicateLog(ds.Tables[0].Rows[i]["docnum"].ToString(), CompanyCode).Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        if (flag == true)
                        {
                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "UPDATE syspex_ePayment Set Status = '1'  where DocNum ='" + ds.Tables[0].Rows[i]["docnum"].ToString() + "' and Company ='" + CompanyCode + "'";
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();
                        }
                        else
                        {
                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "UPDATE syspex_ePayment Set Status = '0'  where DocNum ='" + ds.Tables[0].Rows[i]["docnum"].ToString() + "' and Company ='" + CompanyCode + "'";
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();
                        }

                    }
                    else
                    {

                        if (flag == true)
                        {

                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = @"INSERT INTO syspex_ePayment(Company,DocNum, CustomerCode,CustomerName,ToEmail,Path,SendDate,DocDate,Time,Status,CC) 
                            VALUES(@param1,@param2,@param4,@param5 ,@param6,@param7,@param8,@param9,@param10,@param11,@param12)";
                            cmd.Parameters.AddWithValue("@param1", CompanyCode);
                            cmd.Parameters.AddWithValue("@param2", ds.Tables[0].Rows[i]["docnum"].ToString());
                            cmd.Parameters.AddWithValue("@param4", ds.Tables[0].Rows[i]["CardCode"].ToString());
                            cmd.Parameters.AddWithValue("@param5", ds.Tables[0].Rows[i]["cardname"].ToString());
                            cmd.Parameters.AddWithValue("@param6", ds.Tables[0].Rows[i]["E_Mail"].ToString());
                            cmd.Parameters.AddWithValue("@param7", "F:\\ePayment\\" + CompanyCode + "\\" + ds.Tables[0].Rows[i]["docnum"].ToString() + ".pdf");
                            cmd.Parameters.AddWithValue("@param8", DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
                            cmd.Parameters.AddWithValue("@param9", ds.Tables[0].Rows[i]["docdate"].ToString());
                            cmd.Parameters.AddWithValue("@param10", DateTime.Parse(DateTime.Now.TimeOfDay.ToString()));
                            cmd.Parameters.AddWithValue("@param11", "1");
                            cmd.Parameters.AddWithValue("@param12", ds.Tables[0].Rows[i]["CC"].ToString());
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();

                        }
                        else
                        {
                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = @"INSERT INTO syspex_ePayment(Company,DocNum, CustomerCode,CustomerName,ToEmail,Path,SendDate,DocDate,Time,Status,CC) 
                            VALUES(@param1,@param2,@param4,@param5 ,@param6,@param7,@param8,@param9,@param10,@param11,@param12)";
                            cmd.Parameters.AddWithValue("@param1", CompanyCode);
                            cmd.Parameters.AddWithValue("@param2", ds.Tables[0].Rows[i]["docnum"].ToString());
                            cmd.Parameters.AddWithValue("@param4", ds.Tables[0].Rows[i]["CardCode"].ToString());
                            cmd.Parameters.AddWithValue("@param5", ds.Tables[0].Rows[i]["cardname"].ToString());
                            cmd.Parameters.AddWithValue("@param6", ds.Tables[0].Rows[i]["E_Mail"].ToString());
                            cmd.Parameters.AddWithValue("@param7", "F:\\ePayment\\" + CompanyCode + "\\" + ds.Tables[0].Rows[i]["docnum"].ToString() + ".pdf");
                            cmd.Parameters.AddWithValue("@param8", DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
                            cmd.Parameters.AddWithValue("@param9", ds.Tables[0].Rows[i]["docdate"].ToString());
                            cmd.Parameters.AddWithValue("@param10", DateTime.Parse(DateTime.Now.TimeOfDay.ToString()));
                            cmd.Parameters.AddWithValue("@param11", "0");
                            cmd.Parameters.AddWithValue("@param12", ds.Tables[0].Rows[i]["CC"].ToString());
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();
                        }
                    }
                }
        }

        private static DataSet CheckDuplicateLog(string Docnum, string CompanyCode)
        {
            if (SAPCon12.State == ConnectionState.Closed) { SAPCon12.Open(); }
            DataSet dsetItem = new DataSet();
            SqlCommand CmdItem = new SqlCommand("select DocNum from syspex_ePayment where DocNum ='" + Docnum + "' and Company ='" + CompanyCode + "'", SAPCon12)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SAPCon12.Close();
            return dsetItem;
        }

        private static bool SendInvociePDF(string DocNum, string DocEntry, string To, string CompanyCode, string VendorName)
        {
            bool success;
            string Databasename = "";
            string CC = "";
            //To = "vigna@syspex.com,pheng.teoh@syspex.com,";

            if (CompanyCode == "65ST")
                Databasename = "SYSPEX_LIVE";
            if (CompanyCode == "03SM")
                Databasename = "Syspex Mechatronic (M) Sdn Bhd";
            if (CompanyCode == "07ST")
                Databasename = "Syspex Technologies (M) Sdn Bhd";
            if (CompanyCode == "21SK")
                Databasename = "PT SYSPEX KEMASINDO";
            if (CompanyCode == "31SM")
                Databasename = "PT SYSPEX MULTITECH";
            if (CompanyCode == "04SI")
                Databasename = "Syspex Industries (M) Sdn Bhd";

            try
            {

                ReportDocument cryRpt = new ReportDocument();


                if (CompanyCode == "65ST")
                {
                    cryRpt.Load("F:\\Crystal Reports\\65ST_Payment_Advice.rpt");
                    CC = GetCC("65st", "e_payment");
                }


                if (CompanyCode == "07ST")
                {
                    cryRpt.Load("F:\\Crystal Reports\\07ST_Payment_Advice.rpt");
                    CC = GetCC("07st", "e_payment");
                }

                new TableLogOnInfos();
                TableLogOnInfo crtableLogoninfo;
                var crConnectionInfo = new ConnectionInfo();

                ParameterFieldDefinitions crParameterFieldDefinitions;
                ParameterFieldDefinition crParameterFieldDefinition;
                ParameterValues crParameterValues = new ParameterValues();
                ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

                crParameterDiscreteValue.Value = Convert.ToString(DocEntry);
                crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
                crParameterFieldDefinition = crParameterFieldDefinitions["Dockey@"];
                crParameterValues = crParameterFieldDefinition.CurrentValues;

                crParameterValues.Clear();
                crParameterValues.Add(crParameterDiscreteValue);
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

                crConnectionInfo.ServerName = "SYSPEXSAP04";
                crConnectionInfo.DatabaseName = Databasename;
                crConnectionInfo.UserID = "sa";
                crConnectionInfo.Password = "Password1111";

                var crTables = cryRpt.Database.Tables;
                foreach (Table crTable in crTables)
                {
                    crtableLogoninfo = crTable.LogOnInfo;
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                    crTable.ApplyLogOnInfo(crtableLogoninfo);
                }



                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();

                CrDiskFileDestinationOptions.DiskFileName = "F:\\ePayment\\" + CompanyCode + "\\" + DocNum + ".pdf";
                CrExportOptions = cryRpt.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                cryRpt.Export();

                //// Email Part 

                MailMessage mm = new MailMessage
                {
                    From = new MailAddress("noreply@syspex.com")
                };

                mm.IsBodyHtml = true;
                if (CompanyCode == "65ST")
                {
                    mm.Subject = "Remittance Advice From Syspex Technologies Pte Ltd - " + DocNum + "";
                    mm.Body = "<p>Dear Valued Supplier,</p> <p>This is to inform you payment has been made for the invoices as per attached. Payment will be received in 2 – 3 working days upon receiving this payment advice.</p>" +
     "<p> Regards,</p>" +
"<p>Syspex Technologies Pte Ltd</p> ";
                }
                if (CompanyCode == "07ST")
                {
                    mm.Subject = "Remittance Advice From Syspex Technologies (M) Sdn Bhd - " + DocNum + "";
                    mm.Body = "<p>Dear Valued Supplier,</p> <p>This is to inform you payment has been made for the invoices as per attached. Payment will be received in 2 – 3 working days upon receiving this payment advice.</p>" +
         "<p> Regards,</p>" +
"<p>Syspex Technologies (M) Sdn Bhd</p> ";
                }



                //To
                foreach (var address in To.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (IsValidEmail(address) == true)
                    {
                        mm.To.Add(address);
                    }
                }
                //CC
                foreach (var address in CC.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).Distinct())
                {
                    mm.CC.Add(new MailAddress(address)); //Adding Multiple CC email Id
                }


                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    EnableSsl = true
                };
                System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential
                {
                    UserName = "noreply@syspex.com",
                    Password = "design35"
                };
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = 587;
                mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                smtp.Send(mm);
                success = true;


            }
            catch (CrystalReportsException ex)
            {

                throw ex;
            }

            return success;
        }

        private static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;

            }
            catch
            {
                return false;
            }
        }



        private static DataSet GetOutNos(string CompanyCode)
        {

            SqlConnection SQLConnection = new SqlConnection();

            if (CompanyCode == "65ST")
            {
                SQLConnection = SGConnection;
                SQLQuery = "select top 10 * from (select  docdate,docnum,docentry,CreateDate,CardName,CardCode,  (SELECT STUFF((select ',' + E_MailL from OCPR where CardCode = T0.CardCode  and Name like '%Accounts Receivable%'  FOR XML PATH('')), 1, 1, '')) as E_Mail from ovpm T0 where T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].syspex_ePayment where company ='" + CompanyCode + "')  and year(CreateDate) = year(getdate()) and month(CreateDate) = month(getdate()) and CreateDate <= getdate() ) X where  X.E_Mail is not null order  by X.CreateDate desc";
            }
            if (CompanyCode == "07ST")
            {
                SQLConnection = JBConnection;
                // 15-06-2021 i have chnaged the docdate greater than or equal todays date
                SQLQuery = "select top 10 * from (select  docdate,docnum,docentry,CreateDate,CardName,CardCode,  (SELECT STUFF((select ',' + E_Mail from OCRD where CardCode = T0.CardCode   FOR XML PATH('')), 1, 1, '')) as E_Mail from ovpm T0 where T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].syspex_ePayment where Company='" + CompanyCode + "')    and year(docdate) = year(getdate()) and DocDate >='20210622' and  DocDate <=getdate() ) X where  X.E_Mail is not null order  by X.DocDate desc";
            }
            DataSet dsetItem = new DataSet();
            SqlCommand CmdItem = new SqlCommand(SQLQuery, SQLConnection)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SQLConnection.Close();
            return dsetItem;
        }
        #region ----GetCCFROMPORTAL---
        private static string GetCC(string Branch, string Category)
        {
            if (PortalCon.State == ConnectionState.Closed) { PortalCon.Open(); }
            MySqlCommand CmdItem = new MySqlCommand("SELECT emails FROM sap_email_notification WHERE branch='" + Branch + "' AND category= '" + Category + "' LIMIT 1 ", PortalCon);
            string retrievedValue = "";
            CmdItem.CommandType = CommandType.Text;
            using (MySqlDataReader reader = CmdItem.ExecuteReader())
            {


                while (reader.Read())
                {
                    retrievedValue = (string)reader.GetValue(0);

                }
                reader.Close();
            }
            return retrievedValue;
        }
        #endregion
    }
}
