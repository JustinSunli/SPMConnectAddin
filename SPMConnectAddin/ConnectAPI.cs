using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace SPMConnectAddin
{
    public class ConnectAPI
    {
        private SqlConnection cn;
        private ISldWorks swApp = null;
        private bool doneshowingSplash = false;

        public ConnectAPI(ISldWorks sldworks)
        {
            this.swApp = sldworks;
            SPM_Connect();
        }

        private void SPM_Connect()
        {
            string connection = "Data Source=spm-sql;Initial Catalog=SPM_Database;User ID=SPM_Agent;password=spm5445";
            try
            {
                cn = new SqlConnection(connection);
            }
            catch (Exception)
            {
                MessageBox.Show("Error Connecting to SQL Server.....", "SPM Connect Sql Commands", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string UserName()
        {
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            if (userName.Length > 0)
            {
                return userName;
            }
            else
            {
                return null;
            }
        }

        public string Getassyversionnumber()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string version = "V" + assembly.GetName().Version.ToString(3);
            return version;
        }

        public string Getuserfullname()
        {
            string fullname = "";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [SPM_Database].[dbo].[Users] WHERE [UserName]='" + UserName().ToString() + "' ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    fullname = dr["Name"].ToString();
                }
                dt.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Unable to retrieve user full name", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return fullname;
        }

        public string Getdepartment()
        {
            string Department = "";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [SPM_Database].[dbo].[Users] WHERE [UserName]='" + UserName().ToString() + "' ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    Department = dr["Department"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Unable to retrieve user department", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return Department;
        }

        public void Deleteitem(string _itemno)
        {
            DialogResult result = MessageBox.Show("Are you sure want to delete " + _itemno + "?", "SPM Connect - Delete Item?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                if (_itemno.Length > 0)
                {
                    if (cn.State == ConnectionState.Closed)
                        cn.Open();
                    try
                    {
                        string query = "DELETE FROM [SPM_Database].[dbo].[Inventory] WHERE ItemNumber ='" + _itemno.ToString() + "'";
                        SqlCommand sda = new SqlCommand(query, cn);
                        sda.ExecuteNonQuery();
                        cn.Close();
                        MessageBox.Show(_itemno + " - Is removed from the system now!", "SPM Connect - Delete Item", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "SPM Connect - Delete Item", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cn.Close();
                    }
                }
            }
        }

        public string Getsharesfolder()
        {
            string path = "";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [SPM_Database].[dbo].[Users] WHERE [UserName]='" + UserName() + "' ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    path = dr["SharesFolder"].ToString();
                }
                dt.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Error Getting share folder path", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return path;
        }

        #region UserRights

        public bool CheckAdmin()
        {
            bool admin = false;
            string useradmin = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

            using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[Users] WHERE UserName = @username AND Admin = '1'", cn))
            {
                try
                {
                    cn.Open();
                    sqlCommand.Parameters.AddWithValue("@username", useradmin);

                    int userCount = (int)sqlCommand.ExecuteScalar();
                    if (userCount == 1)
                    {
                        admin = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Unable to retrieve admin rights", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
            }
            return admin;
        }

        public bool Checkdeveloper()
        {
            bool developer = false;
            string useradmin = UserName();

            using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[Users] WHERE UserName = @username AND Developer = '1'", cn))
            {
                try
                {
                    cn.Open();
                    sqlCommand.Parameters.AddWithValue("@username", useradmin);

                    int userCount = (int)sqlCommand.ExecuteScalar();
                    if (userCount == 1)
                    {
                        developer = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Unable to retrieve developer rights", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
            }
            return developer;
        }

        public bool CheckManagement()
        {
            bool management = false;
            using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[Users] WHERE UserName = @username AND Management = '1'", cn))
            {
                try
                {
                    cn.Open();
                    sqlCommand.Parameters.AddWithValue("@username", UserName());

                    int userCount = (int)sqlCommand.ExecuteScalar();
                    if (userCount == 1)
                    {
                        management = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Check management rights", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
            }
            return management;
        }

        public bool CheckingDrawingRights()
        {
            bool management = false;
            using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[Users] WHERE UserName = @username AND [CheckDrawing] = '1'", cn))
            {
                try
                {
                    cn.Open();
                    sqlCommand.Parameters.AddWithValue("@username", UserName());

                    int userCount = (int)sqlCommand.ExecuteScalar();
                    if (userCount == 1)
                    {
                        management = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Checking Drawing rights", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
            }
            return management;
        }

        public bool ApproveDrawingRights()
        {
            bool management = false;
            using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[Users] WHERE UserName = @username AND [ApproveDrawing] = '1'", cn))
            {
                try
                {
                    cn.Open();
                    sqlCommand.Parameters.AddWithValue("@username", UserName());

                    int userCount = (int)sqlCommand.ExecuteScalar();
                    if (userCount == 1)
                    {
                        management = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Approve Drawing rights", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
            }
            return management;
        }

        #endregion UserRights

        public void Chekin(string applicationname)
        {
            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("dd-MM-yyyy HH:mm tt");
            string computername = System.Environment.MachineName;

            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO [SPM_Database].[dbo].[Checkin] ([Last Login],[Application Running],[User Name], [Computer Name], [Version]) VALUES('" + sqlFormattedDate + "', '" + applicationname + "', '" + UserName() + "', '" + computername + "','" + Getassyversionnumber() + "')";
                cmd.ExecuteNonQuery();
                cn.Close();
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - User Checkin", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        public void Checkout()
        {
            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                string query = "DELETE FROM [SPM_Database].[dbo].[Checkin] WHERE [User Name] ='" + UserName().ToString() + "'";
                SqlCommand sda = new SqlCommand(query, cn);
                sda.ExecuteNonQuery();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Checkout User", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        public string Getfilename()
        {
            ModelDoc2 swModel;

            string filename = "";

            int count;
            count = swApp.GetDocumentCount();

            if (count > 0)
            {
                // MessageBox.Show("Number of open documents in this SOLIDWORKS session: " + count);
                swModel = swApp.ActiveDoc as ModelDoc2;

                filename = swModel.GetTitle().Substring(0, 6);
            }
            return filename;
        }

        public string Get_pathname()
        {
            ModelDoc2 swModel;

            int count;
            string pathName = "";
            count = swApp.GetDocumentCount();

            if (count > 0)
            {
                swModel = swApp.ActiveDoc as ModelDoc2;

                pathName = swModel.GetPathName();
            }
            return pathName;
        }

        public string Getfamilycategory(string familycode)
        {
            string category = "";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Category FROM [SPM_Database].[dbo].[FamilyCodes] WHERE [FamilyCodes]='" + familycode.ToString() + "' ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    category = dr["Category"].ToString();
                    //MessageBox.Show(category);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Get Family Category", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return category;
        }

        public string Makepath(string itemnumber)
        {
            string Pathpart = "";

            if (itemnumber.Length > 0)
            {
                string first3char = itemnumber.Substring(0, 3) + @"\";
                string spmcadpath = @"\\spm-adfs\CAD Data\AAACAD\";
                Pathpart = (spmcadpath + first3char);
            }
            return Pathpart;
        }

        public bool Checkforreadonly()
        {
            bool notreadonly = true;

            swApp.Visible = true;
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            if (swModel.IsOpenedReadOnly())
            {
                MessageBox.Show("Model is open read only. Please get write access from the associated user in order to edit the properties.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                notreadonly = false;
            }

            return notreadonly;
        }

        private bool CreateNewItem(string uniqueid)
        {
            bool success = false;
            string user = Getuserfullname();
            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("yyyy-MM-dd HH:mm:ss.fff");
            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO [SPM_Database].[dbo].[Inventory] (ItemNumber, DesignedBy, DateCreated, LastSavedBy, LastEdited, JobPlanning) VALUES('" + uniqueid.ToString() + "','" + user + "','" + sqlFormattedDate + "','" + user + "','" + sqlFormattedDate + "','1')";
                cmd.ExecuteNonQuery();
                cn.Close();
                success = true;
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                success = false;
                MessageBox.Show(ex.Message, "SPM Connect - Create Entry On SQL Inventory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return success;
        }

        private bool UpdateItemToSQL(SPMItem item)
        {
            bool success = false;
            DateTime dateedited = DateTime.Now;
            string user = Getuserfullname();
            string sqlFormattedDate = dateedited.ToString("yyyy-MM-dd HH:mm:ss.fff");
            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE [SPM_Database].[dbo].[Inventory] SET Description = '" + item.Description + "',FamilyCode = '" + item.Family + "',Manufacturer = '" + item.Manufacturer + "',ManufacturerItemNumber = '" + item.ManufacturerItemNo + "',Material = '" + item.Material + "',Spare = '" + item.Spare + "',FamilyType = '" + item.FamilyType + "',SurfaceProtection = '" + item.SurfaceProtection + "',HeatTreatment = '" + item.HeatTreatment + "',LastSavedBy = '" + user + "',Rupture = '" + item.Rupture + "',Notes = '" + item.Notes + "',LastEdited = '" + sqlFormattedDate + "' WHERE ItemNumber = '" + item.ItemNo + "' ";
                cmd.ExecuteNonQuery();
                cn.Close();
                success = true;
                //MessageBox.Show("Item sucessfully saved SPM Connect Server.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                success = false;
                MessageBox.Show(ex.Message, "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return success;
            //perfromlockdown();
        }

        #region GetNewItemNumber or copy items

        public string Getactiveblock()
        {
            string useractiveblock = "";

            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [SPM_Database].[dbo].[Users] where UserName ='" + UserName().ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                foreach (DataRow dr in dt.Rows)
                {
                    useractiveblock = dr["ActiveBlockNumber"].ToString();
                    if (useractiveblock == "")
                    {
                        MessageBox.Show("User has not been assigned a block number. Please contact the admin.", "SPM Connect - Get Active Block Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Get User Active Block", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
            }
            finally
            {
                cn.Close();
            }

            return useractiveblock;
        }

        public string Getlastnumber()
        {
            string blocknumber = Getactiveblock().ToString();

            if (blocknumber == "")
            {
                return "";
            }
            else
            {
                string lastnumber = "";
                try
                {
                    if (cn.State == ConnectionState.Closed)
                        cn.Open();
                    SqlCommand cmd = cn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT MAX(RIGHT(ItemNumber,5)) AS " + blocknumber.ToString() + " FROM [SPM_Database].[dbo].[UnionInventory] WHERE ItemNumber like '" + blocknumber.ToString() + "%' AND LEN(ItemNumber)=6";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        lastnumber = dr[blocknumber].ToString();
                        if (lastnumber == "")
                        {
                            lastnumber = blocknumber.Substring(1) + "000";
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Get Last Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
                return lastnumber;
            }
        }

        public bool CheckBaseBlockNumberTaken()
        {
            string blocknumber = Getactiveblock().ToString();
            bool taken = false;
            if (blocknumber == "")
            {
                return taken;
            }
            else
            {
                string lastnumber = "";
                try
                {
                    if (cn.State == ConnectionState.Closed)
                        cn.Open();
                    SqlCommand cmd = cn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT MAX(RIGHT(ItemNumber,5)) AS " + blocknumber.ToString() + " FROM [SPM_Database].[dbo].[UnionInventory] WHERE ItemNumber like '" + blocknumber.ToString() + "%' AND LEN(ItemNumber)=6";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        lastnumber = dr[blocknumber].ToString();
                        if (lastnumber == "")
                        {
                            taken = false;
                        }
                        else if (lastnumber == blocknumber.Substring(1) + "000")
                        {
                            taken = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Get Last Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
                return taken;
            }
        }

        public string Spmnew_idincrement(string lastnumber, string blocknumber)
        {
            if (!CheckBaseBlockNumberTaken() && lastnumber.Substring(2) == "000")
            {
                string lastnumbergrp1 = blocknumber.Substring(0, 1).ToUpper();
                string newid1 = lastnumbergrp1 + lastnumber.ToString();
                return newid1;
            }
            else
            {
                string lastnumbergrp = blocknumber.Substring(0, 1).ToUpper();
                int lastnumbers = Convert.ToInt32(lastnumber);
                lastnumbers += 1;
                string newid = lastnumbergrp + lastnumbers.ToString();
                return newid;
            }
        }

        public bool Validnumber(string lastnumber)
        {
            bool valid = true;
            if (lastnumber.ToString() != "")
            {
                if (lastnumber.Substring(2) == "999")
                {
                    MessageBox.Show("User block number limit has reached. Please ask the admin to assign a new block number.", "SPM Connect - Valid Number Limit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    valid = false;
                }
            }
            else
            {
                valid = false;
            }
            return valid;
        }

        public bool Checkitempresentoninventory(string itemid)
        {
            bool itempresent = false;
            using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[Inventory] WHERE [ItemNumber]='" + itemid.ToString() + "'", cn))
            {
                try
                {
                    cn.Open();

                    int userCount = (int)sqlCommand.ExecuteScalar();
                    if (userCount == 1)
                    {
                        //MessageBox.Show("item already exists");
                        itempresent = true;
                    }
                    else
                    {
                        //MessageBox.Show(" move forward");
                        itempresent = false;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "SPM Connect - Check Item Present On SQL Inventory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cn.Close();
                }
            }
            return itempresent;
        }

        public void Addcpoieditemtosqltablefromgenius(string newid, string activeid)
        {
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO [SPM_Database].[dbo].[Inventory] (ItemNumber,Description,FamilyCode,Manufacturer,ManufacturerItemNumber,DesignedBy,DateCreated,LastSavedBy,LastEdited) SELECT '" + newid + "',Description,FamilyCode,Manufacturer,ManufacturerItemNumber,DesignedBy,DateCreated,LastSavedBy,LastEdited FROM [SPM_Database].[dbo].[UnionInventory] WHERE ItemNumber = '" + activeid + "'";
                cmd.ExecuteNonQuery();
                cn.Close();
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Add Copied Item To Inventory From Genius", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        public void Addcpoieditemtosqltable(string selecteditem, string uniqueid)
        {
            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("yyyy-MM-dd HH:mm:ss.fff");

            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO [SPM_Database].[dbo].[Inventory](ItemNumber,Description,FamilyCode,Manufacturer,ManufacturerItemNumber,Material,Spare,DesignedBy,FamilyType,SurfaceProtection,HeatTreatment,Rupture,JobPlanning,Notes,DateCreated) SELECT '" + uniqueid + "',Description,FamilyCode,Manufacturer,ManufacturerItemNumber,Material,Spare,DesignedBy,FamilyType,SurfaceProtection,HeatTreatment,Rupture,JobPlanning,Notes,'" + sqlFormattedDate + "' FROM [SPM_Database].[dbo].[Inventory] WHERE ItemNumber = '" + selecteditem + "'";
                cmd.ExecuteNonQuery();
                cn.Close();
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Add Copied Item To Inventory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        #endregion GetNewItemNumber or copy items

        #region OpenModel & Drawing

        public void Checkforspmfile(string Item_No, bool readOnly)
        {
            string ItemNumbero;
            ItemNumbero = Item_No + "-0";

            if (!String.IsNullOrWhiteSpace(Item_No) && Item_No.Length == 6)
            {
                string first3char = Item_No.Substring(0, 3) + @"\";
                //MessageBox.Show(first3char);

                string spmcadpath = @"\\spm-adfs\CAD Data\AAACAD\";

                string Pathpart = (spmcadpath + first3char + Item_No + ".sldprt");
                string Pathassy = (spmcadpath + first3char + Item_No + ".sldasm");
                string PathPartNo = (spmcadpath + first3char + ItemNumbero + ".sldprt");
                string PathAssyNo = (spmcadpath + first3char + ItemNumbero + ".sldasm");

                if (File.Exists(Pathassy) && File.Exists(Pathpart))
                {
                    MessageBox.Show($"System has found a Part file and Assembly file with the same PartNo." + Item_No + "." +
                        " So please contact the administrator.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(PathAssyNo) && File.Exists(PathPartNo))
                {
                    MessageBox.Show($"System has found a Part file and Assembly file with the same PartNo. " + ItemNumbero + "." +
                        " So please contact the administrator.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(PathAssyNo) && File.Exists(Pathpart))
                {
                    MessageBox.Show($"System has found a Part file " + Item_No + "and Assembly file " + ItemNumbero + " with the same PartNo." +
                        " So please contact the administrator.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(Pathassy) && File.Exists(PathPartNo))
                {
                    MessageBox.Show($"System has found a Part file " + ItemNumbero + "and Assembly file" + Item_No + " with the same PartNo." +
                        " So please contact the administrator.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(PathPartNo) && File.Exists(Pathpart))
                {
                    MessageBox.Show($"System has found a Part two files " + Item_No + "," + ItemNumbero + " with the same PartNo." +
                        " So please contact the administrator.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(PathAssyNo) && File.Exists(Pathassy))
                {
                    MessageBox.Show($"System has found a assembly files " + Item_No + "," + ItemNumbero + " with the same PartNo." +
                        " So please contact the administrator.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (File.Exists(Pathassy))
                {
                    Open_assy(Pathassy, readOnly);
                }
                else if (File.Exists(PathAssyNo))
                {
                    Open_assy(PathAssyNo, readOnly);
                }
                else if (File.Exists(Pathpart))
                {
                    Open_model(Pathpart, readOnly);
                }
                else if (File.Exists(PathPartNo))
                {
                    Open_model(PathPartNo, readOnly);
                }
                else
                {
                    MessageBox.Show($"A file with the part number " + Item_No + " does not have Solidworks CAD Model. Please Try Again.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //fName = "";
                }
            }
            doneshowingSplash = true;
        }

        public void Open_model(string filename, bool readOnly)
        {
            swApp.Visible = true;
            int err = 0;
            int warn = 0;
            ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc6(filename, (int)swDocumentTypes_e.swDocPART, readOnly ? ((int)swOpenDocOptions_e.swOpenDocOptions_LoadLightweight + (int)swOpenDocOptions_e.swOpenDocOptions_Silent + (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly) : (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
            swApp.ActivateDoc(filename);
            swModel = swApp.ActiveDoc as ModelDoc2;
        }

        public void Open_assy(string filename, bool readOnly)
        {
            swApp.Visible = true;
            int err = 0;
            int warn = 0;
            ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc6(filename, (int)swDocumentTypes_e.swDocASSEMBLY, readOnly ? ((int)swOpenDocOptions_e.swOpenDocOptions_LoadLightweight + (int)swOpenDocOptions_e.swOpenDocOptions_Silent + (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly) : (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
            swApp.ActivateDoc(filename);
            swModel = swApp.ActiveDoc as ModelDoc2;
        }

        public void Open_drw(string filename)
        {
            swApp.Visible = true;
            int err = 0;
            int warn = 0;
            ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc6(filename, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
            swApp.ActivateDoc(filename);
            swModel = swApp.ActiveDoc as ModelDoc2;
        }

        #endregion OpenModel & Drawing

        #region solidworks createmodels and open models

        public void Createmodel(string filename)
        {
            swApp.Visible = true;
            string PartPath = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(PartPath, 0, 0, 0) as ModelDoc2;
            swApp.Visible = true;
            swModel = swApp.ActiveDoc as ModelDoc2;
            ModelDocExtension swExt;
            swExt = swModel.Extension;
            bool boolstatus = false;
            boolstatus = swExt.SaveAs(filename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            swApp.ActivateDoc(filename);
            swModel = swApp.ActiveDoc as ModelDoc2;

            if (boolstatus == true)
            {
                //MessageBox.Show("new part created");
                Get_pathname();
                Getfilename();
            }
            else
            {
                //MessageBox.Show("part not saved");
            }
        }

        public void Createassy(string filename)
        {
            swApp.Visible = true;
            string assytemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            ModelDoc2 swModel = swApp.NewDocument(assytemplate, 0, 0, 0) as ModelDoc2;
            swApp.Visible = true;
            swModel = swApp.ActiveDoc as ModelDoc2;
            ModelDocExtension swExt;
            swExt = swModel.Extension;
            bool boolstatus = false;
            //boolstatus = swExt.SaveAs(filename, 0, (int)swDocumentTypes_e.swDocASSEMBLY, 0, 0, 0);
            boolstatus = swExt.SaveAs(filename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            swApp.ActivateDoc(filename);
            swModel = swApp.ActiveDoc as ModelDoc2;

            if (boolstatus == true)
            {
                //MessageBox.Show("new assy created");
                Get_pathname();
                Getfilename();
            }
            else
            {
                //MessageBox.Show("part not saved");
            }
        }

        public bool Createdrawingpart(string filename, string _itemnumber)
        {
            swApp.Visible = true;
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
            ModelDoc2 swModel;
            DrawingDoc swDrawing;
            ModelDocExtension swModelDocExt;

            swModel = (ModelDoc2)swApp.NewDocument(template, (int)swDwgPaperSizes_e.swDwgPaperBsize, 0, 0);
            swDrawing = (DrawingDoc)swModel;
            swDrawing = swApp.ActiveDoc as DrawingDoc;
            swModelDocExt = (ModelDocExtension)swModel.Extension;

            string Pathpart = Makepath(_itemnumber).ToString() + _itemnumber + ".sldprt";

            swDrawing.Create3rdAngleViews2(Pathpart);

            //Sheet cursheet;
            //cursheet = swDrawing.GetCurrentSheet();
            //double sheetwidth = 0;
            //double sheethieght = 0;
            //int lRetVal;
            //lRetVal= cursheet.GetSize(sheetwidth,sheethieght);
            //SolidWorks.Interop.sldworks.View swView;

            //swView = (SolidWorks.Interop.sldworks.View)swDrawing.CreateDrawViewFromModelView3(Pathpart, "*Isometric",sheetwidth, sheethieght, 0);
            //swDrawing.InsertModelAnnotations3(0, 327663, true, true, false, false);
            //int nNumView = 0;
            //var Voutline;
            //var Vpostion;
            //double viewweidth = 0;
            //double viewheight = 0;

            //Voutline(nNumView) = swView.GetOutline();
            //Vpostion(nNumView) = swView.Position();
            //viewweidth = Voutline(2) - Voutline(0);

            //swView.Position(6, 5);

            bool boolstatus = false;
            boolstatus = swModelDocExt.SaveAs(filename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            swApp.QuitDoc(swModel.GetTitle());

            return boolstatus;
        }

        public bool Createdrawingassy(string filename, string _itemnumber)
        {
            swApp.Visible = true;
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
            ModelDoc2 swModel;
            DrawingDoc swDrawing;
            ModelDocExtension swModelDocExt;

            swModel = (ModelDoc2)swApp.NewDocument(template, (int)swDwgPaperSizes_e.swDwgPaperBsize, 0, 0);
            swDrawing = (DrawingDoc)swModel;
            swDrawing = swApp.ActiveDoc as DrawingDoc;
            swModelDocExt = (ModelDocExtension)swModel.Extension;

            string Pathpart = Makepath(_itemnumber).ToString() + _itemnumber + ".sldasm";

            swDrawing.Create3rdAngleViews2(Pathpart);

            //Sheet cursheet;
            //cursheet = swDrawing.GetCurrentSheet();
            //double sheetwidth = 0;
            //double sheethieght = 0;
            //int lRetVal;
            //lRetVal= cursheet.GetSize(sheetwidth,sheethieght);
            //SolidWorks.Interop.sldworks.View swView;

            //swView = (SolidWorks.Interop.sldworks.View)swDrawing.CreateDrawViewFromModelView3(Pathpart, "*Isometric",sheetwidth, sheethieght, 0);
            //swDrawing.InsertModelAnnotations3(0, 327663, true, true, false, false);
            //int nNumView = 0;
            //var Voutline;
            //var Vpostion;
            //double viewweidth = 0;
            //double viewheight = 0;

            //Voutline(nNumView) = swView.GetOutline();
            //Vpostion(nNumView) = swView.Position();
            //viewweidth = Voutline(2) - Voutline(0);

            //swView.Position(6, 5);

            bool boolstatus = false;
            boolstatus = swModelDocExt.SaveAs(filename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            swApp.QuitDoc(swModel.GetTitle());

            return boolstatus;
        }

        public bool Importstepfile(string stepFileName, string savefilename)
        {
            PartDoc swPart = default(PartDoc);
            AssemblyDoc swassy = default(AssemblyDoc);
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            ImportStepData swImportStepData = default(ImportStepData);

            bool status = false;
            int errors = 0;

            //Get import information
            swImportStepData = (ImportStepData)swApp.GetImportFileData(stepFileName);

            //If ImportStepData::MapConfigurationData is not set, then default to
            //the environment setting swImportStepConfigData; otherwise, override
            //swImportStepConfigData with ImportStepData::MapConfigurationData
            swImportStepData.MapConfigurationData = true;

            //Import the STEP file.
            try
            {
                swPart = (PartDoc)swApp.LoadFile4(stepFileName, "r", swImportStepData, ref errors);
                swModel = (ModelDoc2)swPart;
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                status = swModelDocExt.SaveAs(savefilename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            }
            catch
            {
            }
            try
            {
                swassy = (AssemblyDoc)swApp.LoadFile4(stepFileName, "r", swImportStepData, ref errors);
                swModel = (ModelDoc2)swPart;
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                status = swModelDocExt.SaveAs(savefilename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            }
            catch
            {
            }

            return status;
        }

        public bool Importigesfile(string igesfilename, string savefilename)
        {
            PartDoc swPart = default(PartDoc);
            AssemblyDoc swassy = default(AssemblyDoc);
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            ImportIgesData swImportIgesdata = default(ImportIgesData);

            bool status = false;
            int errors = 0;
            swImportIgesdata = (ImportIgesData)swApp.GetImportFileData(igesfilename);
            swImportIgesdata.IncludeSurfaces = true;
            try
            {
                swPart = (PartDoc)swApp.LoadFile4(igesfilename, "r", swImportIgesdata, ref errors);
                swModel = (ModelDoc2)swPart;
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                status = swModelDocExt.SaveAs(savefilename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            }
            catch
            {
            }
            try
            {
                swassy = (AssemblyDoc)swApp.LoadFile4(igesfilename, "r", swImportIgesdata, ref errors);
                swModel = (ModelDoc2)swassy;
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                status = swModelDocExt.SaveAs(savefilename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            }
            catch
            {
            }

            return status;
        }

        public bool Importparasolidfile(string parasolidfilename, string savefilename)
        {
            PartDoc swPart = default(PartDoc);
            AssemblyDoc swassy = default(AssemblyDoc);
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);

            bool status = false;
            int errors = 0;
            try
            {
                swPart = (PartDoc)swApp.LoadFile4(parasolidfilename, "r", null, ref errors);
                swModel = (ModelDoc2)swPart;
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                status = swModelDocExt.SaveAs(savefilename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            }
            catch
            {
            }
            try
            {
                swassy = (AssemblyDoc)swApp.LoadFile4(parasolidfilename, "r", null, ref errors);
                swModel = (ModelDoc2)swassy;
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                status = swModelDocExt.SaveAs(savefilename, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0, 0);
            }
            catch
            {
            }

            return status;
        }

        #endregion solidworks createmodels and open models

        #region Copy

        private string GetItemFamily(string itemnumber)
        {
            string category = "";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT *  FROM [SPM_Database].[dbo].[Inventory] WHERE [ItemNumber]='" + itemnumber.ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    category = dr["FamilyCode"].ToString();
                    //MessageBox.Show(category);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Get Item Family Code", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return category;
        }

        public void Prepareforcopy(string activeblock, string selecteditem, string lastnumber)
        {
            string first3char = selecteditem.Substring(0, 3) + @"\";
            string spmcadpath = @"\\spm-adfs\CAD Data\AAACAD\";
            string Pathpart = (spmcadpath + first3char);

            if (lastnumber.ToString().Length > 0)
            {
                string uniqueid = Spmnew_idincrement(lastnumber.ToString(), activeblock.ToString());

                if (Checkitempresentoninventory(uniqueid) == true)
                {
                    //insertinto_blocks(uniqueid, activeblock.ToString());
                    MessageBox.Show("SPM Item number already exixts with your new part number.", "SPM Connect - Copy Model", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Checkforspmfile(uniqueid, false);
                }
                else
                {
                    bool sucessreplacingreference = Copy(Pathpart, selecteditem, uniqueid);
                    Aftercopy(activeblock, selecteditem, uniqueid, sucessreplacingreference);
                    FillData(uniqueid);
                }
            }
            else
            {
                string uniqueid = Spmnew_id(activeblock.ToString());
                bool sucessreplacingreference = Copy(Pathpart, selecteditem, uniqueid);
                Aftercopy(activeblock, selecteditem, uniqueid, sucessreplacingreference);
                FillData(uniqueid);
            }
        }

        private string Spmnew_id(string blocknumber)
        {
            string letterblock = Char.ToUpper(blocknumber[0]) + blocknumber.Substring(1);
            return letterblock + "000";
        }

        private void Aftercopy(string activeblock, string selecteditem, string uniqueid, bool sucessreplacingreference)
        {
            if (sucessreplacingreference)
            {
                if (Checkitempresentoninventory(selecteditem))
                {
                    Addcpoieditemtosqltable(selecteditem, uniqueid);
                }
                else
                {
                    Addcpoieditemtosqltablefromgenius(uniqueid, selecteditem);
                    Updateitemtosqlinventory(uniqueid);
                }

                Checkforspmfile(uniqueid, false);
            }
            else
            {
                MessageBox.Show("SPM Connect failed to update drawing references.! Please manually update drawing references.", "SPM Connect - Copy References", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (Checkitempresentoninventory(selecteditem))
                {
                    Addcpoieditemtosqltable(selecteditem, uniqueid);
                }
                else
                {
                    Addcpoieditemtosqltablefromgenius(uniqueid, selecteditem);
                    Updateitemtosqlinventory(uniqueid);
                }
                Checkforspmfile(uniqueid, false);
            }
        }

        private void Updateitemtosqlinventory(string uniqueid)
        {
            string familycategory = Getfamilycategory(GetItemFamily(uniqueid).ToString());
            //MessageBox.Show(familycategory);
            string rupture = "ALWAYS";

            if (familycategory.ToLower() == "purchased")
            {
                rupture = "NEVER";
            }
            string username = Getuserfullname();

            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("yyyy-MM-dd HH:mm:ss.fff");
            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE [SPM_Database].[dbo].[Inventory] SET FamilyType = '" + familycategory + "',Rupture = '" + rupture + "',JobPlanning = '1',LastSavedBy = '" + username + "',DateCreated = '" + sqlFormattedDate + "',LastEdited = '" + sqlFormattedDate + "'  WHERE ItemNumber = '" + uniqueid + "'";
                cmd.ExecuteNonQuery();
                cn.Close();
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Update Item SQL Inventory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        private bool Copy(string Pathpart, string selecteditem, string uniqueid)
        {
            string type = "";
            string drawingfound = "no";
            string oldpath = "";
            string newfirst3char = uniqueid.Substring(0, 3) + @"\";
            string spmcadpath = @"\\spm-adfs\CAD Data\AAACAD\";
            bool success = false;

            string[] s = Directory.GetFiles(Pathpart, "*" + selecteditem + "*", SearchOption.TopDirectoryOnly).Where(str => !str.Contains(@"\~$")).ToArray();

            for (int i = 0; i < s.Length; i++)
            {
                //MessageBox.Show(s[i]);
                //MessageBox.Show(Path.GetFileName(s[i]));

                if (s[i].ToLower().Contains(".sldprt"))
                {
                    //MessageBox.Show("found part");
                    type = "part";
                    oldpath = s[i];
                }
                else if (s[i].ToLower().Contains(".sldasm"))
                {
                    //MessageBox.Show("found assy");
                    type = "assy";
                    oldpath = s[i];
                }
                else if (s[i].ToLower().Contains(".slddrw"))
                {
                    //MessageBox.Show("found assy");
                    drawingfound = "yes";
                }
                string filename = Path.GetFileName(s[i]);
                string extension = filename.Substring(filename.IndexOf('.'));

                string newfilepathdir = spmcadpath + newfirst3char;
                System.IO.Directory.CreateDirectory(newfilepathdir);

                string newfileexits = spmcadpath + newfirst3char + uniqueid + extension;

                if (File.Exists(newfileexits))
                {
                    if (MessageBox.Show(newfileexits + " already exists\r\nDo you want to overwrite it?", "Overwrite File - SPM Connect - Copy File Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                    {
                        File.Copy(s[i], newfileexits, true);
                    }
                    else
                    {
                        return success;
                    }
                }
                else
                {
                    File.Copy(s[i], newfileexits, false);
                }
            }

            if (drawingfound == "yes")
            {
                string newdraw = spmcadpath + newfirst3char + uniqueid + ".slddrw";
                string newpath = "";
                if (type == "part")
                {
                    newpath = spmcadpath + newfirst3char + uniqueid + ".sldprt";
                }
                else if (type == "assy")
                {
                    newpath = spmcadpath + newfirst3char + uniqueid + ".sldasm";
                }

                success = Replacereference(newdraw, oldpath, newpath);
            }
            else
            {
                success = true;
            }
            return success;
        }

        private bool Replacereference(string newdraw, string oldpath, string newpath)
        {
            return swApp.ReplaceReferencedDocument(newdraw, oldpath, newpath);
        }

        private void FillData(string itemnumber)
        {
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT *  FROM [SPM_Database].[dbo].[Inventory] WHERE [ItemNumber]='" + itemnumber.ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                List<string> list = Graballinfor(dt);
                Chekbeforefillingcustomproperties(itemnumber, list);
                MessageBox.Show("File successfully copied to new item number.", "SPM Connect - Copy Model", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Get Family Category", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        private List<string> Graballinfor(DataTable dt)
        {
            List<string> list = new List<string>();
            list.Clear();
            DataRow r = dt.Rows[0];
            string itemnumber = r["ItemNumber"].ToString();
            string description = r["Description"].ToString();
            string manufacturer = r["Manufacturer"].ToString();
            string oem = r["ManufacturerItemNumber"].ToString();
            string material = r["Material"].ToString();
            string designby = r["DesignedBy"].ToString();
            string familytype = r["FamilyType"].ToString();
            string surface = r["SurfaceProtection"].ToString();
            string heat = r["HeatTreatment"].ToString();
            string datecreated = r["DateCreated"].ToString();
            string dateedit = r["LastEdited"].ToString();
            string lastsaved = r["LastSavedBy"].ToString();
            string notes = r["Notes"].ToString();
            string rupture = r["Rupture"].ToString();
            string family = r["FamilyCode"].ToString();
            string spare = r["Spare"].ToString();

            if (family.Length > 0)
            {
            }
            else
            {
                family = "MA";
                familytype = "Manufactured";
                rupture = "ALWAYS";
            }

            list.Add(itemnumber);
            list.Add(description);
            list.Add(material);
            list.Add(manufacturer);
            list.Add(oem);
            list.Add(family);
            list.Add(familytype);
            list.Add(surface);
            list.Add(heat);
            list.Add(notes);
            list.Add(rupture);
            list.Add(spare);
            list.Add(designby);
            list.Add(datecreated);
            list.Add(lastsaved);
            list.Add(dateedit);

            for (int i = 0; i < list.Count; i++)
            {
                list[i] = list[i].Replace("'", "''");
            }
            return list;
        }

        public void Chekbeforefillingcustomproperties(string item, List<string> list)
        {
            string getcurrentfilename = Getfilename().ToString();
            string olditemnumber = item + "-0";
            if (getcurrentfilename == item || getcurrentfilename == olditemnumber)
            {
                Fillcustomproperties(list);
            }
            else
            {
                //if (checkforreadonly() == true)
                //{
                //    fillcustomproperties();
                //}

                string Pathassy = Makepath(item).ToString() + item + ".sldasm";
                string Pathpart = Makepath(item).ToString() + item + ".sldprt";
                string Pathassyo = Makepath(item).ToString() + item + "-0" + ".sldasm";
                string Pathparto = Makepath(item).ToString() + item + "-0" + ".sldprt";

                if (File.Exists(Pathassy))
                {
                    Open_assy(Pathassy, false);
                    Fillcustomproperties(list);
                }
                else if (File.Exists(Pathpart))
                {
                    Open_model(Pathpart, false);
                    Fillcustomproperties(list);
                }
                else if (File.Exists(Pathparto))
                {
                    Open_model(Pathparto, false);
                    Fillcustomproperties(list);
                }
                else if (File.Exists(Pathassyo))
                {
                    Open_assy(Pathassyo, false);
                    Fillcustomproperties(list);
                }
                else
                {
                    MessageBox.Show("Please have the active model open in order to save custom properties to the soliworks document..", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        public void Fillcustomproperties(List<string> list)
        {
            try
            {
                var progId = "SldWorks.Application";

                SldWorks swApp = System.Runtime.InteropServices.Marshal.GetActiveObject(progId.ToString()) as SolidWorks.Interop.sldworks.SldWorks;
                ModelDoc2 swModel;
                CustomPropertyManager cusPropMgr;
                int lRetVal;
                swModel = (ModelDoc2)swApp.ActiveDoc;
                ModelDocExtension swModelDocExt = default(ModelDocExtension);
                swModelDocExt = swModel.Extension;
                cusPropMgr = swModelDocExt.get_CustomPropertyManager("");
                lRetVal = cusPropMgr.Add3("PartNo", (int)swCustomInfoType_e.swCustomInfoText, list[0].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Description", (int)swCustomInfoType_e.swCustomInfoText, list[1].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("OEM", (int)swCustomInfoType_e.swCustomInfoText, list[3].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("OEM Item Number", (int)swCustomInfoType_e.swCustomInfoText, list[4].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("cDesignedBy", (int)swCustomInfoType_e.swCustomInfoText, list[12].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Heat Treatment", (int)swCustomInfoType_e.swCustomInfoText, list[8].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Surface Protection", (int)swCustomInfoType_e.swCustomInfoText, list[7].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Spare", (int)swCustomInfoType_e.swCustomInfoText, list[11].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("JobPlanning", (int)swCustomInfoType_e.swCustomInfoText, "1", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Notes", (int)swCustomInfoType_e.swCustomInfoText, list[9].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Rupture", (int)swCustomInfoType_e.swCustomInfoText, list[10].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Heat Treatment Req'd", (int)swCustomInfoType_e.swCustomInfoText, list[8].ToString().Length > 0 ? "Checked" : "Unchecked", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Surface Protection Req'd", (int)swCustomInfoType_e.swCustomInfoText, list[7].ToString().Length > 0 ? "Checked" : "Unchecked", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("Family Type", (int)swCustomInfoType_e.swCustomInfoText, list[6].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                lRetVal = cusPropMgr.Add3("cCategory", (int)swCustomInfoType_e.swCustomInfoText, list[5].ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                string category = Getfamilycategory(list[5].ToString()).ToString();
                if (category.ToLower() == "manufactured")
                {
                    PartDoc swPart = default(PartDoc);
                    swPart = (PartDoc)swModel;
                    swPart.SetMaterialPropertyName2("Default", "//SPM-ADFS/CAD Data/CAD Templates SPM/SPM.sldmat", list[2].ToString());
                }

                bool boolstatus = false;
                boolstatus = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect New Item Fill Custom Properties", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion Copy

        private void SplashDialog(string message)
        {
            doneshowingSplash = false;
            System.Threading.ThreadPool.QueueUserWorkItem((x) =>
            {
                using (var splashForm = new Dialog())
                {
                    splashForm.TopMost = true;
                    splashForm.Focus();
                    splashForm.Activate();
                    splashForm.Message = message;
                    splashForm.StartPosition = FormStartPosition.CenterScreen;
                    splashForm.Show();
                    while (!doneshowingSplash)
                        Application.DoEvents();
                    splashForm.Close();
                }
            });
        }

        #region Favorites

        public bool Addtofavorites(string itemid)
        {
            if (itemid == "")
            {
                itemid = Getfilename();
            }
            bool completed = false;
            if (ValidfileName(itemid))
            {
                if (CheckitempresentonFavorites(itemid))
                {
                    string usernamesfromitem = Getusernamesfromfavorites(itemid);
                    if (!Userexists(usernamesfromitem))
                    {
                        string newuseradded = usernamesfromitem + UserName() + ",";
                        Updateusernametoitemonfavorites(itemid, newuseradded);
                    }
                    else
                    {
                        MessageBox.Show("Item no " + itemid + " already exists on your favorite list.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    Additemtofavoritessql(itemid);
                }
            }
            else
            {
                MessageBox.Show($"A file with the part number " + itemid + " does not have Solidworks CAD Model or SPM item number assigned. Cannot add or remove from favorites. Please Try Again.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return completed;
        }

        public bool Removefromfavorites(string itemid)
        {
            if (itemid == "")
            {
                itemid = Getfilename();
            }
            bool completed = false;
            if (ValidfileName(itemid))
            {
                string usernamesfromitem = Getusernamesfromfavorites(itemid);

                Updateusernametoitemonfavorites(itemid, Removeusername(usernamesfromitem));

                MessageBox.Show("Item no " + itemid + " has been removed from your favorite list.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"A file with the part number " + itemid + " does not have Solidworks CAD Model or SPM item number assigned. Cannot add or remove from favorites. Please Try Again.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return completed;
        }

        public bool CheckitempresentonFavorites(string itemid)
        {
            bool itempresent = false;
            if (itemid == "")
            {
                itemid = Getfilename();
            }
            if (ValidfileName(itemid))
            {
                using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[favourite] WHERE [Item]='" + itemid.ToString() + "'", cn))
                {
                    try
                    {
                        cn.Open();

                        int userCount = (int)sqlCommand.ExecuteScalar();
                        if (userCount == 1)
                        {
                            //MessageBox.Show("item already exists");
                            itempresent = true;
                        }
                        else
                        {
                            //MessageBox.Show(" move forward");
                            itempresent = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "SPM Connect - Check Item Present On SQL Favorites", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cn.Close();
                    }
                }
            }
            return itempresent;
        }

        private bool ValidfileName(string Item_No)
        {
            bool validitem = true;
            string ItemNumbero;
            ItemNumbero = Item_No + "-0";

            if (!String.IsNullOrWhiteSpace(Item_No) && Item_No.Length == 6)
            {
                string first3char = Item_No.Substring(0, 3) + @"\";
                //MessageBox.Show(first3char);

                string spmcadpath = @"\\spm-adfs\CAD Data\AAACAD\";

                string Pathpart = (spmcadpath + first3char + Item_No + ".sldprt");
                string Pathassy = (spmcadpath + first3char + Item_No + ".sldasm");
                string PathPartNo = (spmcadpath + first3char + ItemNumbero + ".sldprt");
                string PathAssyNo = (spmcadpath + first3char + ItemNumbero + ".sldasm");

                if (File.Exists(Pathassy) && File.Exists(Pathpart))
                {
                }
                else if (File.Exists(PathAssyNo) && File.Exists(PathPartNo))
                {
                }
                else if (File.Exists(PathAssyNo) && File.Exists(Pathpart))
                {
                }
                else if (File.Exists(Pathassy) && File.Exists(PathPartNo))
                {
                }
                else if (File.Exists(PathPartNo) && File.Exists(Pathpart))
                {
                }
                else if (File.Exists(PathAssyNo) && File.Exists(Pathassy))
                {
                }
                else if (File.Exists(Pathassy))
                {
                }
                else if (File.Exists(PathAssyNo))
                {
                }
                else if (File.Exists(Pathpart))
                {
                }
                else if (File.Exists(PathPartNo))
                {
                }
                else
                {
                    validitem = false;
                }
            }
            else
            {
                validitem = false;
            }
            return validitem;
        }

        private void Additemtofavoritessql(string itemid)
        {
            string userid = UserName();
            userid += ",";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO [SPM_Database].[dbo].[favourite] (Item,UserName) VALUES('" + itemid + "','" + userid + " ')";
                cmd.ExecuteNonQuery();
                cn.Close();
                MessageBox.Show("Item no " + itemid + " has been added to your favorites.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Add  Item To Favorites", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        private void Updateusernametoitemonfavorites(string itemid, string updatedusername)
        {
            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                if (updatedusername != "")
                {
                    cmd.CommandText = "UPDATE [SPM_Database].[dbo].[favourite] SET UserName = '" + updatedusername + "'  WHERE Item = '" + itemid + "'";
                }
                else
                {
                    cmd.CommandText = "DELETE FROM [SPM_Database].[dbo].[favourite] WHERE Item = '" + itemid + "'";
                }

                cmd.ExecuteNonQuery();
                cn.Close();
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Update Item on Favorites", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        private string Getusernamesfromfavorites(string itemid)
        {
            string usersfav = "";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [SPM_Database].[dbo].[favourite] WHERE [Item]='" + itemid + "' ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    usersfav = dr["UserName"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Unable to retrieve user names from favorites", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return usersfav;
        }

        private bool Userexists(string usernames)
        {
            bool exists = false;
            string userid = UserName();
            // Split string on spaces (this will separate all the words).
            string[] words = usernames.Split(',');
            foreach (string word in words)
            {
                if (word == userid)
                    exists = true;
            }

            return exists;
        }

        private string Removeusername(string usernames)
        {
            string removedusername = "";
            string userid = UserName();
            // Split string on spaces (this will separate all the words).
            string[] words = usernames.Split(',');
            foreach (string word in words)
            {
                if (word.Trim() == userid)
                {
                }
                else
                {
                    removedusername += word.Trim();
                    if (word.Trim() != "")
                        removedusername += ",";
                }
            }
            return removedusername.Trim();
        }

        #endregion Favorites

        public void Randomcolor()
        {
            swApp.Visible = true;
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                // Tell user
                MessageBox.Show("Active model is not a part", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            try
            {
                //SelectionMgr swSelMgr = default(SelectionMgr);
                //swSelMgr = (SelectionMgr)swModel.SelectionManager;
                //Face2 swFace = default(Face2);
                //int count = swSelMgr.GetSelectedObjectCount2(-1);
                //for (int i = 1; i < count; i++)
                //{
                //    swFace = swSelMgr.GetSelectedObject6(i, -1);
                //    var featColors = swFace.MaterialPropertyValues;
                //    int highTrim = 382;
                //    Random rnd = new Random();
                //    Byte[] b = new Byte[3];
                //    rnd.NextBytes(b);
                //    Color color = Color.FromArgb(b[0], b[1], b[2]);
                //    if ((color.R + color.G + color.B) > highTrim)
                //    {
                //        color = Color.FromArgb(255 - color.R, 255 - color.G, 255 - color.B);
                //    }
                //    featColors[0] = color.R;
                //    featColors[1] = color.G;
                //    featColors[2] = color.B;
                //    swFace.MaterialPropertyValues(featColors);
                //}

                SelectionMgr swSelMgr = default(SelectionMgr);
                SelectData swSelData = default(SelectData);
                Feature swFeat = default(Feature);
                double[] featColors = null;
                Random rnd = new Random();
                int highTrim = 382;
                Byte[] b = new Byte[3];
                rnd.NextBytes(b);
                Color color = Color.FromArgb(b[0], b[1], b[2]);
                if ((color.R + color.G + color.B) > highTrim)
                {
                    color = Color.FromArgb(255 - color.R, 255 - color.G, 255 - color.B);
                }
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swSelMgr = (SelectionMgr)swModel.SelectionManager;

                int count = swSelMgr.GetSelectedObjectCount2(-1);
                if (count > 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        swFeat = (Feature)swSelMgr.GetSelectedObject6(i + 1, -1);
                        swSelData = (SelectData)swSelMgr.CreateSelectData();
                        featColors = (double[])swModel.MaterialPropertyValues;
                        featColors[0] = color.R;
                        featColors[1] = color.G;
                        featColors[2] = color.B;
                        swFeat.SetMaterialPropertyValues(featColors);

                        //faceArr = (object[])swFeat.GetFaces();
                        //if ((faceArr == null))
                        //    return;
                        //foreach (object oneFace in faceArr)
                        //{
                        //    swFace = (Face2)oneFace;
                        //    swEnt = (Entity)swFace;
                        //    swFaceFeat = (Feature)swFace.GetFeature();
                        //    // Check to see if face is owned by multiple features
                        //    if (object.ReferenceEquals(swFaceFeat, swFeat))
                        //    {
                        //        status = swEnt.Select4(true, swSelData);

                        //        swFace.MaterialPropertyValues = featColors;
                        //    }
                        //    else
                        //    {
                        //    }
                        //}
                    }
                }
                else
                {
                    var vMatProp = swModel.MaterialPropertyValues;
                    vMatProp[0] = color.R;
                    vMatProp[1] = color.G;
                    vMatProp[2] = color.B;
                    swModel.MaterialPropertyValues = vMatProp;
                }

                swModel.ClearSelection2(true);
                swModel.EditRebuild3();
                swModel.ViewZoomtofit2();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Random Color", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CreateCube()
        {
            try
            {
                //make sure we have a part open
                string partTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
                if ((partTemplate != null) && (partTemplate != ""))
                {
                    IModelDoc2 swDoc = (IModelDoc2)swApp.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);

                    // Create sketch
                    swDoc.SketchManager.InsertSketch(false);
                    object[] rectangle;
                    rectangle = (object[])swDoc.SketchManager.CreateCenterRectangle(0.0, 0.0, 0.0, 0.05, 0.025, 0.0);
                    SketchSegment l1, l2;
                    l1 = (SketchSegment)rectangle[0];
                    l2 = (SketchSegment)rectangle[1];
                    l1.Select4(false, null);
                    swDoc.AddDimension2(0.0, 0.075, 0.0);
                    l2.Select4(false, null);
                    swDoc.AddDimension2(0.075, 0.0125, 0.0);
                    swDoc.SketchManager.InsertSketch(true);
                    //Extrude the sketch
                    IFeatureManager featMan = swDoc.FeatureManager;
                    featMan.FeatureExtrusion(true,
                        false, false,
                        (int)swEndConditions_e.swEndCondMidPlane, (int)swEndConditions_e.swEndCondMidPlane,
                        0.1, 0.0,
                        false, false,
                        false, false,
                        0.0, 0.0,
                        false, false,
                        false, false,
                        true,
                        false, false);
                    swDoc.ViewZoomtofit2();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Create Cube", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InsertM6FPT()
        {
            try
            {
                //make sure we have a part open
                ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
                if (swModel == null)
                {
                    MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    // Tell user
                    MessageBox.Show("Active model is not a part", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                // Create sketch
                swModel.SketchManager.InsertSketch(true);
                swModel.SketchManager.MakeSketchBlockFromFile(null, @"\\spm-adfs\CAD Data\AAACAD\A52\A52572-0.SLDBLK", false, 1.0, 0.0);
                swModel.SetPickMode();
                swModel.ClearSelection2(true);
                swModel.ViewZoomtofit2();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Create Cube", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ReloadSheetformat()
        {
            string sTemplatePath = @"\\spm-adfs\CAD Data\CAD Templates SPM\";
            string sTemplateName = "GENIUS B - landscape.slddrt";
            string sTemplateNameD = "GENIUS D - landscape.slddrt";

            ModelDoc2 swModel = default(ModelDoc2);
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                // Tell user
                MessageBox.Show("Active model is not a drawing", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DrawingDoc swDrawDoc = default(DrawingDoc);
            swDrawDoc = (DrawingDoc)swModel;
            Sheet swSheet = default(Sheet);

            string[] obj = null;

            obj = (string[])swDrawDoc.GetSheetNames();
            swSheet = swDrawDoc.GetCurrentSheet();

            string gttemplatename = "";
            try
            {
                foreach (string vName in obj)
                {
                    swDrawDoc.ActivateSheet(vName);
                    swSheet = (Sheet)swDrawDoc.GetCurrentSheet();
                    gttemplatename = swSheet.GetTemplateName();
                    gttemplatename = gttemplatename.Substring(gttemplatename.Length - 27);
                    var objs = swSheet.GetProperties();

                    if (gttemplatename == sTemplateName || gttemplatename == "genius b - landscape.slddrt" || gttemplatename == "pm\\aaa b - landscape.slddrt")
                    {
                        swDrawDoc.SetupSheet5(swSheet.GetName(), (int)swDwgPaperSizes_e.swDwgPapersUserDefined, (int)swDwgTemplates_e.swDwgTemplateNone, (double)objs[2], (double)objs[3], false, "", 0.4318, 0.2794, "Default", true);

                        swDrawDoc.SetupSheet5(swSheet.GetName(), (int)swDwgPaperSizes_e.swDwgPapersUserDefined, (int)swDwgTemplates_e.swDwgTemplateCustom, (double)objs[2], (double)objs[3], false, sTemplatePath + sTemplateName, 0.4318, 0.2794, "Default", true);
                        swModel.ViewZoomtofit2();
                    }
                    else
                    {
                        swDrawDoc.SetupSheet5(swSheet.GetName(), (int)swDwgPaperSizes_e.swDwgPapersUserDefined, (int)swDwgTemplates_e.swDwgTemplateNone, (double)objs[2], (double)objs[3], false, "", 0.4318, 0.2794, "Default", true);
                        swDrawDoc.SetupSheet5(swSheet.GetName(), (int)swDwgPaperSizes_e.swDwgPapersUserDefined, (int)swDwgTemplates_e.swDwgTemplateCustom, (double)objs[2], (double)objs[3], false, sTemplatePath + sTemplateNameD, 0.4318, 0.2794, "Default", true);

                        swModel.ViewZoomtofit2();
                    }
                }
                swDrawDoc.ActivateSheet(obj[0]);
                swModel.ForceRebuild3(false);
                //swModel.Save3(1, 0, 0);
                // MessageBox.Show("Successfully reloaded sheet format.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void DeleteDanglingDimensions()
        {
            ModelDoc2 swModel = default(ModelDoc2);
            swModel = (ModelDoc2)swApp.ActiveDoc;
            bool boolstatus = false;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                // Tell user
                MessageBox.Show("Active model is not a drawing", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DrawingDoc swDrawDoc = default(DrawingDoc);
            swDrawDoc = (DrawingDoc)swModel;
            Sheet swSheet = default(Sheet);
            SolidWorks.Interop.sldworks.View swView = default(SolidWorks.Interop.sldworks.View);
            SolidWorks.Interop.sldworks.Annotation swAnn = default(SolidWorks.Interop.sldworks.Annotation);

            swModel.ClearSelection2(true);
            string[] vSheetNames = null;

            vSheetNames = (string[])swDrawDoc.GetSheetNames();
            swSheet = swDrawDoc.GetCurrentSheet();
            foreach (string vName in vSheetNames)
            {
                swDrawDoc.ActivateSheet(vName);
                swSheet = (Sheet)swDrawDoc.GetCurrentSheet();
                swView = swDrawDoc.GetFirstView();
                while (swView != null)
                {
                    swAnn = swView.GetFirstAnnotation3();
                    while (swAnn != null)
                    {
                        if (swAnn.IsDangling())
                            boolstatus = swAnn.Select3(true, null);
                        swAnn = swAnn.GetNext3();
                    }

                    swView = swView.GetNextView();
                    boolstatus = swModel.DeleteSelection(true);
                    swModel.ClearSelection2(true);
                }
            }
            swModel.ClearSelection2(true);
            if (boolstatus == false)
                swApp.SendMsgToUser("Failed to fix dangling dimensions.");
            else
                // Tell user
                MessageBox.Show("Successfully deleted dangling dimensions. If there are any missing dangling dimensions, run the program again.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void CloseInactive()
        {
            try
            {
                ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
                if (swModel == null)
                {
                    MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                ModelDoc2 swModelref = swApp.ActiveDoc as ModelDoc2;
                string swModelName = swModel.GetTitle();
                int swDocType = swModel.GetType();
                object[] vModels = (object[])swApp.GetDocuments();
                int docCount = 0;
                List<String> list = new List<String>();
                int count = swApp.GetDocumentCount();
                for (int i = 0; i < count; i++)
                {
                    swModelref = vModels[i] as ModelDoc2;
                    if (swModelref.GetTitle() != swModelName)
                    {
                        list.Add(swModelref.GetTitle());
                        docCount++;
                    }
                }

                for (int i = 0; i < docCount; i++)
                {
                    swApp.CloseDoc(list[i]);
                }
                MessageBox.Show("Successfully closed inactive documents.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception)
            {
            }
        }

        public void GetMyNextPartNo()
        {
            if (Validnumber(Getlastnumber()))
            {
                string newid = Spmnew_idincrement(Getlastnumber(), Getactiveblock().ToString());
                MessageBox.Show("Your next ItemNumber to use is :- " + newid, "SPM Connect - New Item Number", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Clipboard.SetText(newid);
            }
        }

        #region Export Features

        /// <summary>
        /// Exports the currently active part as a DXF
        /// </summary>
        public void ExportPartAsDxf()
        {
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                // Tell user
                MessageBox.Show("Active model is not a part", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                return;
            }
            PartDoc swPart = (PartDoc)swModel;
            string filepath = Getsharesfolder();
            if (string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show("User shares folder is not configured. Please contact admin", "SPM Connect - Share Folder Path", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            filepath = filepath + @"\SPM_Connect\DXF\";
            System.IO.Directory.CreateDirectory(filepath);
            filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + ".DXF";

            SelectionMgr swSelMgr = default(SelectionMgr);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            int count = swSelMgr.GetSelectedObjectCount2(-1);

            if (count == 1) // one face selected
            {
                //Check for 2d face
                int seltype = swSelMgr.GetSelectedObjectType3(1, 0);
                if (seltype != (int)swSelectType_e.swSelFACES)
                {
                    MessageBox.Show("Please select a (one) 2D face before running this command.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    bool isThisAPlane = false;
                    if (swSelMgr.GetSelectedObjectType3(1, -1) != 2)
                    {
                        MessageBox.Show("Item selected is not a face. You must select a (one) complete 2D Face to export.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        Face2 WorkFace = default(Face2);
                        WorkFace = swSelMgr.GetSelectedObject5(1);
                        Surface WorkSurface = default(Surface);
                        WorkSurface = WorkFace.IGetSurface();
                        isThisAPlane = WorkSurface.IsPlane();
                        if (isThisAPlane == false)
                        {
                            MessageBox.Show("Item selected is not a face. You must select a (one) complete 2D Face to export.");
                            return;
                        }
                        else
                        {
                            object varAlignment;
                            double[] dataAlignment = new double[12];
                            string[] dataViews = new string[2];
                            dataAlignment[0] = 0.0;
                            dataAlignment[1] = 0.0;
                            dataAlignment[2] = 0.0;
                            dataAlignment[3] = 1.0;
                            dataAlignment[4] = 0.0;
                            dataAlignment[5] = 0.0;
                            dataAlignment[6] = 0.0;
                            dataAlignment[7] = 1.0;
                            dataAlignment[8] = 0.0;
                            dataAlignment[9] = 0.0;
                            dataAlignment[10] = 0.0;
                            dataAlignment[11] = 1.0;
                            varAlignment = dataAlignment;
                            bool retVal = swPart.ExportToDWG2(filepath, swModel.GetPathName(), (int)swExportToDWG_e.swExportToDWG_ExportSelectedFacesOrLoops, true, varAlignment, false, false, 0, null);
                            if (retVal == false)
                                MessageBox.Show("Failed to save flat face as DXF", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            else
                            {
                                // Tell user success
                                DialogResult result = MessageBox.Show("Successfully saved flat face as DXF at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                                if (result == DialogResult.Yes)
                                {
                                    string argument = "/select, \"" + filepath + "\"";
                                    System.Diagnostics.Process.Start("explorer.exe", argument);
                                }
                                else
                                {
                                }
                            }
                        }
                    }
                }
            }
            else if (count == 0)
            {
                object varAlignment;
                double[] dataAlignment = new double[12];
                object varViews;
                string[] dataViews = new string[10];

                dataAlignment[0] = 0.0;
                dataAlignment[1] = 0.0;
                dataAlignment[2] = 0.0;
                dataAlignment[3] = 1.0;
                dataAlignment[4] = 0.0;
                dataAlignment[5] = 0.0;
                dataAlignment[6] = 0.0;
                dataAlignment[7] = 1.0;
                dataAlignment[8] = 0.0;
                dataAlignment[9] = 0.0;
                dataAlignment[10] = 0.0;
                dataAlignment[11] = 1.0;

                varAlignment = dataAlignment;

                dataViews[0] = "*Current";
                dataViews[1] = "*Front";
                dataViews[2] = "*Top";
                dataViews[3] = "*Bottom";
                dataViews[4] = "*Right";
                dataViews[5] = "*Left";
                dataViews[6] = "*Back";
                dataViews[7] = "*Isometric";
                dataViews[8] = "*Trimetric";
                dataViews[9] = "*Dimetric";
                varViews = dataViews;
                bool retVal = swPart.ExportToDWG2(filepath, swModel.GetPathName(), (int)swExportToDWG_e.swExportToDWG_ExportAnnotationViews, true, varAlignment, false, false, 0, varViews);
                if (retVal == false)
                    MessageBox.Show("Failed to save model as DXF", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    // Tell user success
                    DialogResult result = MessageBox.Show("Successfully saved model as DXF at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (result == DialogResult.Yes)
                    {
                        string argument = "/select, \"" + filepath + "\"";
                        System.Diagnostics.Process.Start("explorer.exe", argument);
                    }
                    else
                    {
                    }
                }
            }
            else
            {
                // multiple faces selected
                MessageBox.Show("Please select a (one) 2D face before running this command. Or else select zero entities to save the model views as DXF.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// Exports the currently active part as a STEP
        /// </summary>
        public void ExportModelAsStep()
        {
            // Make sure we have a part or assembly
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            bool boolstatus = false;

            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                MessageBox.Show("Active model is not a part or assembly", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            string filepath = Getsharesfolder();
            if (string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show("User shares folder is not configured. Please contact admin", "SPM Connect - Share Folder Path", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            filepath = filepath + @"\SPM_Connect\STEP\";
            System.IO.Directory.CreateDirectory(filepath);
            filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + ".STEP";

            boolstatus = swModel.SaveAs(filepath);

            if (!boolstatus)
                // Tell user failed
                MessageBox.Show("Failed to save model as STEP", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                // Tell user success
                DialogResult result = MessageBox.Show("Successfully saved model as STEP at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    string argument = "/select, \"" + filepath + "\"";
                    System.Diagnostics.Process.Start("explorer.exe", argument);
                }
                else
                {
                }
            }
        }

        /// <summary>
        /// Exports the currently active part as a PDF
        /// </summary>
        public void ExportDrawingAsPdf()
        {
            // Make sure we have a part or assembly
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            bool boolstatus = false;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                // Tell user
                MessageBox.Show("Active model is not a drawing", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            // Ask the user where to export the file
            //var location = GetSaveLocation("PDF File|*.pdf", "Save Part as PDF");

            // If the user canceled, return
            string filepath = Getsharesfolder();
            if (string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show("User shares folder is not configured. Please contact admin", "SPM Connect - Share Folder Path", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            filepath = filepath + @"\SPM_Connect\Pdfs\";
            System.IO.Directory.CreateDirectory(filepath);
            filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + ".pdf";

            boolstatus = swModel.SaveAs(filepath);

            if (!boolstatus)
                // Tell user failed
                MessageBox.Show("Failed to save drawing as PDF", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                // Tell user success
                DialogResult result = MessageBox.Show("Successfully saved drawing as PDF at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    string argument = "/select, \"" + filepath + "\"";
                    System.Diagnostics.Process.Start("explorer.exe", argument);
                }
                else
                {
                }
            }
        }

        public void ExportModelAsParasolid()
        {
            // Make sure we have a part or assembly
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            bool boolstatus = false;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                MessageBox.Show("Active model is not a part or assembly", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            // Ask the user where to export the file
            //var location = GetSaveLocation("PDF File|*.pdf", "Save Part as PDF");

            // If the user canceled, return
            string filepath = Getsharesfolder();
            if (string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show("User shares folder is not configured. Please contact admin", "SPM Connect - Share Folder Path", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            filepath = filepath + @"\SPM_Connect\Parasolids\";
            System.IO.Directory.CreateDirectory(filepath);
            filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + ".X_T";

            boolstatus = swModel.SaveAs(filepath);

            if (!boolstatus)
                // Tell user failed
                MessageBox.Show("Failed to save model as Parasolid", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                // Tell user success
                DialogResult result = MessageBox.Show("Successfully saved model as Parasolid at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    string argument = "/select, \"" + filepath + "\"";
                    System.Diagnostics.Process.Start("explorer.exe", argument);
                }
                else
                {
                }
            }
        }

        public void ExportModelAsIGES()
        {
            // Make sure we have a part or assembly
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            bool boolstatus = false;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                MessageBox.Show("Active model is not a part or assembly", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            // Ask the user where to export the file
            //var location = GetSaveLocation("PDF File|*.pdf", "Save Part as PDF");

            // If the user canceled, return
            string filepath = Getsharesfolder();
            if (string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show("User shares folder is not configured. Please contact admin", "SPM Connect - Share Folder Path", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            filepath = filepath + @"\SPM_Connect\IGES\";
            System.IO.Directory.CreateDirectory(filepath);
            filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + ".IGS";

            boolstatus = swModel.SaveAs(filepath);

            if (!boolstatus)
                // Tell user failed
                MessageBox.Show("Failed to save model as IGES", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                // Tell user success
                DialogResult result = MessageBox.Show("Successfully saved model as IGES at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    string argument = "/select, \"" + filepath + "\"";
                    System.Diagnostics.Process.Start("explorer.exe", argument);
                }
                else
                {
                }
            }
        }

        public void ExportModelAsIGESToCNC()
        {
            // Make sure we have a part or assembly
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            bool boolstatus = false;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                MessageBox.Show("Active model is not a part or assembly", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            // Ask the user where to export the file
            //var location = GetSaveLocation("PDF File|*.pdf", "Save Part as PDF");

            // If the user canceled, return
            string filepath = @"\\SPM-ADFS\Shares\CNC-Genius\";
            string subdir = swModel.GetTitle().ToString().Substring(0, 3);

            string input = Interaction.InputBox("Please enter the Revison number.", "Revision Number", "", -1, -1);
            if (!(string.IsNullOrEmpty(input)))
            {
                filepath = filepath + subdir + @"\";
                System.IO.Directory.CreateDirectory(filepath);

                filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + " REV-" + input + ".IGS";

                boolstatus = swModel.SaveAs(filepath);

                if (!boolstatus)
                    // Tell user failed
                    MessageBox.Show("Failed to save model as IGES", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    // Tell user success
                    DialogResult result = MessageBox.Show("Successfully saved model as IGES at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (result == DialogResult.Yes)
                    {
                        string argument = "/select, \"" + filepath + "\"";
                        System.Diagnostics.Process.Start("explorer.exe", argument);
                    }
                    else
                    {
                    }
                }
            }
        }

        public void ExportModelAsParasolidToCNC()
        {
            // Make sure we have a part or assembly
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            bool boolstatus = false;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                MessageBox.Show("Active model is not a part or assembly", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            // Ask the user where to export the file
            //var location = GetSaveLocation("PDF File|*.pdf", "Save Part as PDF");

            // If the user canceled, return
            string filepath = @"\\SPM-ADFS\Shares\CNC-Genius\";
            string subdir = swModel.GetTitle().ToString().Substring(0, 3);

            string input = Interaction.InputBox("Please enter the Revison number.", "Revision Number", "", -1, -1);

            if (!(string.IsNullOrEmpty(input)))
            {
                filepath = filepath + subdir + @"\";
                System.IO.Directory.CreateDirectory(filepath);

                filepath = filepath + swModel.GetTitle().ToString().Substring(0, 6) + " REV-" + input + ".X_T";
                boolstatus = swModel.SaveAs(filepath);

                if (!boolstatus)
                    // Tell user failed
                    MessageBox.Show("Failed to save model as Parasolid", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    // Tell user success
                    DialogResult result = MessageBox.Show("Successfully saved model as Parasolid at " + filepath + System.Environment.NewLine + " Would you like to open the file location?", "SPM Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (result == DialogResult.Yes)
                    {
                        string argument = "/select, \"" + filepath + "\"";
                        System.Diagnostics.Process.Start("explorer.exe", argument);
                    }
                    else
                    {
                    }
                }
            }
        }

        #endregion Export Features

        #region Save To Server

        public bool SaveToServer(bool userclick)
        {
            bool success = false;
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                if (userclick)
                    MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                if (userclick)
                    MessageBox.Show("Active model is not a part or assembly", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                return false;
            }

            SPMItem _item = GetActivePartCustomProps(swModel, userclick);
            if (_item == null || _item.ItemNo.Length == 0)
            {
                if (userclick)
                    MessageBox.Show("Part number not found. Please add unique SPM item id to custom properties.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                if (Checkitempresentoninventory(_item.ItemNo) && swModel.GetPathName() != null)
                {
                    // update
                    //grab customproperties
                    if (ValidfileName(_item.ItemNo))// file exists to be updated
                    {
                        if (Getfilename() == _item.ItemNo)
                        {
                            if (UpdateItemToSQL(_item))
                            {
                                if (userclick)
                                    MessageBox.Show("File successfully updated to SPM Connect.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                success = true;
                                return success;
                            }
                            else
                            {
                                if (userclick)
                                    MessageBox.Show("Error updating solidworks custom properties to SPM Connect.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return false;
                            }
                        }
                        else
                        {
                            // filename and partnumber on custom prop doesnt match
                            if (userclick)
                                MessageBox.Show("Active solidworks file and custom properties part number does not match. Cannot update custom properties.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return false;
                        }
                    }
                    else
                    {
                        // there is no file with this partnumber
                        if (userclick)
                            MessageBox.Show("No file present with this partnumber saved on CAD Data.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                }
                else
                {
                    if (!(userclick))
                        return false;
                    // create new item and update
                    if (ValidfileName(_item.ItemNo))//that means file exists
                    {
                        // file with this part number already exists and cannot be saved
                        if (userclick)
                            MessageBox.Show("File with matching partno already exists on CAD Data. Cannot save to server.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                    else
                    {
                        // continue to create new and save
                        // check for oem
                        SPMItem oemexistsItem = CheckOEMExists(_item.ManufacturerItemNo);
                        if (oemexistsItem.ItemNo != null && oemexistsItem.ItemNo.Length > 0)
                        {
                            if (userclick)
                                MessageBox.Show("PartNo Already Exists in Database with the same OEM Item Number - " + _item.ManufacturerItemNo + "." + System.Environment.NewLine + " Use ItemNumber " + oemexistsItem.ItemNo + " - " + oemexistsItem.Description + " (" + oemexistsItem.Manufacturer + ") To Prevent Duplicates.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return false;
                        }
                        else
                        {
                            // Part is valid.. create it
                            // save the file and create a entry on the database
                            DialogResult result = MessageBox.Show("Are you sure want to create a new item?", "SPM Connect - Add New Item?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (result == DialogResult.Yes)
                            {
                                string category = Getfamilycategory(_item.Family).ToString();
                                if (ProcessModelCreation(_item.ItemNo, swModel, category == "Purchased" ? false : true))
                                {
                                    if (CreateNewItem(_item.ItemNo))
                                    {
                                        if (UpdateItemToSQL(_item))
                                        {
                                            if (userclick)
                                                MessageBox.Show("File successfully saved to SPM Connect.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            success = true;
                                            return success;
                                        }
                                        else
                                        {
                                            if (userclick)
                                                MessageBox.Show("Error updating solidworks custom properties to SPM Connect.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        if (userclick)
                                            MessageBox.Show("Error creating new item to SPM Connect.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        return false;
                                    }
                                }
                                else
                                {
                                    // error creating or saving the model
                                    if (userclick)
                                        MessageBox.Show("Error saving or creating solidworks file.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                    return false;
                                }
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }
            }
        }

        private bool ProcessModelCreation(string item, ModelDoc2 swModel, bool needsdrw)
        {
            string path = Makepath(item).ToString();
            bool success = false;
            bool successpart = false;
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                System.IO.Directory.CreateDirectory(path);
                string filename = path + (item) + ".sldprt";
                successpart = swModel.SaveAs(filename);
                if (needsdrw)
                {
                    string draw = path + (item) + ".slddrw";
                    success = Createdrawingpart(draw, item);
                }
                else
                {
                    success = true;
                }
            }
            else if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                System.IO.Directory.CreateDirectory(path);
                string filename = path + (item) + ".sldasm";
                successpart = swModel.SaveAs(filename);
                if (needsdrw)
                {
                    string draw = path + (item) + ".slddrw";
                    success = Createdrawingassy(draw, item);
                }
                else
                {
                    success = true;
                }
            }
            if (success && successpart)
            {
                success = true;
            }
            else
            {
                success = false;
            }
            return success;
        }

        public SPMItem GetActivePartCustomProps(ModelDoc2 swModel, bool userclick)
        {
            // Make sure we have a part or assembly
            SPMItem item = new SPMItem();
            try
            {
                CustomPropertyManager cusPropMgr;
                ModelDocExtension swModelDocExt = default(ModelDocExtension);
                swModelDocExt = swModel.Extension;
                cusPropMgr = swModelDocExt.get_CustomPropertyManager("");

                string partnumber = cusPropMgr.Get("PartNo").Replace("'", "''");

                if (partnumber.Length != 6 && String.IsNullOrEmpty(partnumber) && !Char.IsLetter(partnumber[0]))
                {
                    if (userclick)
                        MessageBox.Show("Active model does not have a valid SPM partnumber", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return null;
                }
                item.ItemNo = partnumber;
                item.Description = cusPropMgr.Get("Description").Replace("'", "''");
                item.Manufacturer = cusPropMgr.Get("OEM").Replace("'", "''");
                item.ManufacturerItemNo = cusPropMgr.Get("OEM Item Number").Replace("'", "''");
                item.Family = cusPropMgr.Get("cCategory").Replace("'", "''");
                item.FamilyType = cusPropMgr.Get("Family Type").Replace("'", "''");
                item.Spare = cusPropMgr.Get("Spare").Replace("'", "''").ToUpper();
                item.SurfaceProtection = cusPropMgr.Get("Surface Protection").Replace("'", "''");
                item.HeatTreatment = cusPropMgr.Get("Heat Treatment").Replace("'", "''");
                item.Rupture = cusPropMgr.Get("Rupture").Replace("'", "''");
                item.Notes = cusPropMgr.Get("Notes").Replace("'", "''");
                cusPropMgr.Get2("cSubCategory", out string textexp, out string materialval);
                item.Material = materialval.Replace("'", "''");
                return item;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Get Custom Properties", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public SPMItem CheckOEMExists(string oemno)
        {
            SPMItem item = new SPMItem();

            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [SPM_Database].[dbo].[Inventory] where (isnull(ManufacturerItemNumber, '') <> '' and ManufacturerItemNumber <> '-') and ManufacturerItemNumber = '" + oemno + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                    foreach (DataRow dr in dt.Rows)
                    {
                        item.ItemNo = dr["ItemNumber"].ToString();
                        item.Description = dr["Description"].ToString();
                        item.Family = dr["FamilyCode"].ToString();
                        item.Manufacturer = dr["Manufacturer"].ToString();
                        item.ManufacturerItemNo = dr["ManufacturerItemNumber"].ToString();
                        break;
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Check OEM Matches", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
            return item;
        }

        #endregion Save To Server

        #region Print and Traverse Components

        public void TraverseFeatureFeatures(Feature swFeat, long nLevel)
        {
            Feature swSubFeat;
            Feature swSubSubFeat;
            Feature swSubSubSubFeat;
            string sPadStr = " ";
            long i = 0;

            for (i = 0; i <= nLevel; i++)
            {
                sPadStr = sPadStr + " ";
            }
            while ((swFeat != null))
            {
                Debug.Print(sPadStr + swFeat.Name + " [" + swFeat.GetTypeName2() + "]");
                swSubFeat = (Feature)swFeat.GetFirstSubFeature();

                while ((swSubFeat != null))
                {
                    Debug.Print(sPadStr + "  " + swSubFeat.Name + " [" + swSubFeat.GetTypeName() + "]");
                    swSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();

                    while ((swSubSubFeat != null))
                    {
                        Debug.Print(sPadStr + "    " + swSubSubFeat.Name + " [" + swSubSubFeat.GetTypeName() + "]");
                        swSubSubSubFeat = (Feature)swSubSubFeat.GetFirstSubFeature();

                        while ((swSubSubSubFeat != null))
                        {
                            Debug.Print(sPadStr + "      " + swSubSubSubFeat.Name + " [" + swSubSubSubFeat.GetTypeName() + "]");
                            swSubSubSubFeat = (Feature)swSubSubSubFeat.GetNextSubFeature();
                        }

                        swSubSubFeat = (Feature)swSubSubFeat.GetNextSubFeature();
                    }

                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
                }

                swFeat = (Feature)swFeat.GetNextFeature();
            }
        }

        public void TraverseComponentFeatures(Component2 swComp, long nLevel)
        {
            Feature swFeat;

            swFeat = (Feature)swComp.FirstFeature();
            TraverseFeatureFeatures(swFeat, nLevel);
        }

        private List<BOM> TraverseComponent(Component2 swComp, long nLevel)
        {
            object[] vChildComp;
            Component2 swChildComp;
            string sPadStr = " ";
            long i = 0;

            for (i = 0; i <= nLevel - 1; i++)
            {
                sPadStr = sPadStr + " ";
            }

            vChildComp = (object[])swComp.GetChildren();
            List<BOM> bomlist = new List<BOM>();
            for (i = 0; i < vChildComp.Length; i++)
            {
                swChildComp = (Component2)vChildComp[i];
                if (swChildComp.GetSuppression() != (int)swComponentSuppressionState_e.swComponentSuppressed && !(swChildComp.ExcludeFromBOM))
                {
                    int bomPos = FindBomPosition(bomlist, swChildComp);
                    if (bomPos == -1)
                    {
                        BOM bOM = new BOM
                        {
                            Path = swChildComp.GetPathName(),
                            Id = bomlist.Count + 1,
                            ItemNo = swChildComp.Name,
                            Description = swChildComp.Name2,
                            Qty = 1
                        };
                        bomlist.Add(bOM);
                    }
                    else
                    {
                        BOM dog = bomlist.FirstOrDefault(d => d.Path.ToLower() == swChildComp.GetPathName().ToLower());
                        if (dog.Path != null) { dog.Qty++; }
                    }

                    Debug.Print(sPadStr + "+" + swChildComp.Name2 + " <" + swChildComp.ReferencedConfiguration + ">");
                }

                //TraverseComponentFeatures(swChildComp, nLevel);
                //TraverseComponent(swChildComp, nLevel + 1);
            }

            foreach (BOM bom in bomlist)
            {
                Debug.Print(bom.Id + "+" + bom.ItemNo + " <" + bom.Description + "> " + bom.Qty);
            }
            return bomlist;
        }

        private int FindBomPosition(List<BOM> bom, Component2 comp)
        {
            foreach (BOM item in bom)
            {
                if ((item.Path).ToLower() == comp.GetPathName().ToLower())
                {
                    return 0;
                }
            }
            return -1;
        }

        public void TraverseModelFeatures(ModelDoc2 swModel, long nLevel)
        {
            Feature swFeat;

            swFeat = (Feature)swModel.FirstFeature();
            TraverseFeatureFeatures(swFeat, nLevel);
        }

        public void ExportBOM()
        {
            ModelDoc2 swModel;
            ConfigurationManager swConfMgr;
            Configuration swConf;
            Component2 swRootComp;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
            swConf = (Configuration)swConfMgr.ActiveConfiguration;
            swRootComp = (Component2)swConf.GetRootComponent();

            System.Diagnostics.Stopwatch myStopwatch = new Stopwatch();
            myStopwatch.Start();

            Debug.Print("File = " + swModel.GetPathName());

            //TraverseModelFeatures(swModel, 1);

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                TraverseComponent(swRootComp, 1);
            }

            myStopwatch.Stop();
            TimeSpan myTimespan = myStopwatch.Elapsed;
            Debug.Print("Time = " + myTimespan.TotalSeconds + " sec");
        }

        public bool IsNumeric(string input)
        {
            int number;
            return int.TryParse(input, out number);
        }

        public void PrintDrawings(Form1 form1)
        {
            int i = 1;
            ModelDoc2 swModel;
            ConfigurationManager swConfMgr;
            Configuration swConf;
            Component2 swRootComp;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                return;
            }

            string job = "";
            string subassy = "";
            bool verfiedjob = false;
            bool verifiedassy = false;

            do
            {
                job = Interaction.InputBox("Enter SPM Job Number", "SPM Connect", "");
                if (job == "")
                {
                    return;
                }
                if (IsNumeric(job) && job.Length == 5)
                {
                    verfiedjob = true;
                }
                else
                {
                    MessageBox.Show("Invalid Job Number. Please try again", "SPM Connect - Job Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } while (verfiedjob == false);

            do
            {
                subassy = Interaction.InputBox("Enter SPM Sub-Assy Number", "SPM Connect", "");
                if (subassy == "")
                {
                    return;
                }

                if (subassy.Length == 6 && !String.IsNullOrEmpty(subassy) && Char.IsLetter(subassy[0]))
                {
                    verifiedassy = true;
                }
                else
                {
                    MessageBox.Show("Not a valid part number. Please enter a valid six digit SPM item number (starting with 'A', 'B', 'C') as valid sub-assy number.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } while (verifiedassy == false);

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                //grab job and subassy no
                string drawing = Checkforspmdrwfile(Getfilename());
                if (drawing != null)
                {
                    form1.UpdateProgressBar(50, "Currenty printing : " + drawing);
                    Open_drw(drawing);
                    ReloadSheetformat();
                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    SendToPrinter(swModel, job, subassy);
                    swApp.QuitDoc(swModel.GetTitle());
                    form1.UpdateProgressBar(100, "Printing Complete. Please collect the print from the printer.");
                }
            }

            //TraverseModelFeatures(swModel, 1);
            List<BOM> bomlist = new List<BOM>();
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
                swConf = (Configuration)swConfMgr.ActiveConfiguration;
                swRootComp = (Component2)swConf.GetRootComponent();

                bomlist = TraverseComponent(swRootComp, 1);

                string assydrw = Checkforspmdrwfile(Getfilename());
                if (assydrw != null)
                {
                    form1.UpdateProgressBar(i, "Currenty printing : " + assydrw);

                    Open_drw(assydrw);
                    ReloadSheetformat();
                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    SendToPrinter(swModel, job, subassy);
                    swApp.QuitDoc(swModel.GetTitle());
                }
                if (bomlist.Count > 0)
                {
                    List<string> drawings = new List<string>();

                    foreach (BOM bom in bomlist)
                    {
                        string drawing = Checkforspmdrwfile(bom.ItemNo);
                        if (drawing != null)
                        {
                            drawings.Add(drawing);
                        }
                    }

                    int incrementby = 100 / drawings.Count;
                    i += incrementby;
                    foreach (string draw in drawings)
                    {
                        form1.UpdateProgressBar(i, "Currenty printing : " + draw);
                        Open_drw(draw);
                        ReloadSheetformat();
                        swModel = (ModelDoc2)swApp.ActiveDoc;
                        SendToPrinter(swModel, job, subassy);

                        swApp.QuitDoc(swModel.GetTitle());
                        i = i + incrementby;
                    }
                }

                form1.UpdateProgressBar(100, "Printing Complete. Please collect the prints from the printer.");
            }
        }

        public string Checkforspmdrwfile(string Item_No)
        {
            Item_No = Item_No.Substring(0, 6);
            string drawing = null;
            string ItemNumbero = Item_No + "-0";

            if (!String.IsNullOrWhiteSpace(Item_No) && Item_No.Length == 6)
            {
                string first3char = Item_No.Substring(0, 3) + @"\";
                //MessageBox.Show(first3char);

                string spmcadpath = @"\\spm-adfs\CAD Data\AAACAD\";

                string Drawpath = (spmcadpath + first3char + Item_No + ".SLDDRW");

                string drawpathno = (spmcadpath + first3char + ItemNumbero + ".SLDDRW");

                if (File.Exists(drawpathno) && File.Exists(Drawpath))
                {
                    return null;
                }
                else if (File.Exists(Drawpath))
                {
                    drawing = Drawpath;
                }
                else if (File.Exists(drawpathno))
                {
                    drawing = drawpathno;
                }
                else
                {
                }
            }
            else
            {
                return null;
            }
            return drawing;
        }

        private void SendToPrinter(ModelDoc2 swModel, string job, string subassy)
        {
            try
            {
                PageSetup ps;
                ModelDocExtension swExt;
                PrintSpecification printSpec;

                swExt = swModel.Extension;
                swExt.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_Document;
                ps = (PageSetup)swModel.PageSetup;
                swModel.DeleteCustomInfo("Job");
                swModel.DeleteCustomInfo("Subassy");
                swModel.AddCustomInfo3("", "Job", (int)swCustomInfoType_e.swCustomInfoText, job);
                swModel.AddCustomInfo3("", "Subassy", (int)swCustomInfoType_e.swCustomInfoText, subassy);
                swModel.ForceRebuild3(true);

                ps.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                ps.ScaleToFit = true;
                swModel.PrintSetup[0] = 17;

                //Create a print specification
                printSpec = (PrintSpecification)swExt.GetPrintSpecification();
                printSpec.ConvertToHighQuality = true;
                printSpec.PrinterQueue = "";
                printSpec.PrintToFile = false;

                swExt.PrintOut4("", "", printSpec);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Send to Printer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //print
        }

        #endregion Print and Traverse Components

        #region Private Helpers

        /// <summary>
        /// Asks the user for a save filename and location
        /// </summary>
        /// <param name="filter">The filter for the save dialog box</param>
        /// <param name="title">The title of the dialog box</param>
        /// <returns></returns>
        private string GetSaveLocation(string filter, string title)
        {
            // Create dialog
            var dialog = new SaveFileDialog { Filter = filter, Title = title, AddExtension = true };

            // Get dialog result
            if (dialog.ShowDialog() == DialogResult.OK)
                return dialog.FileName;
            else
                return null;
        }

        #endregion Private Helpers

        #region Checking and Approving Drawings

        private bool AddToCheckedDrawing(string itemid)
        {
            if (itemid == "")
            {
                itemid = Getfilename();
            }
            bool completed = false;
            if (ValidfileName(itemid))
            {
                if (CheckMarkedDrawingExists(itemid))
                {
                    UpdateDrawingToSQLForChecked(itemid);
                }
                else
                {
                    AddDrawToChecked(itemid);
                }
            }
            else
            {
                //MessageBox.Show($"A file with the part number " + itemid + " does not have Solidworks CAD Model or SPM item number assigned. Cannot add or remove from favorites. Please Try Again.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return completed;
        }

        private bool AddToApprovedDrawing(string itemid)
        {
            if (itemid == "")
            {
                itemid = Getfilename();
            }
            bool completed = false;
            if (ValidfileName(itemid))
            {
                if (CheckMarkedDrawingExists(itemid))
                {
                    UpdateDrawToApproved(itemid);
                }
                else
                {
                    //cannot be marked approved until checked
                }
            }
            else
            {
                //MessageBox.Show($"A file with the part number " + itemid + " does not have Solidworks CAD Model or SPM item number assigned. Cannot add or remove from favorites. Please Try Again.", "SPM-Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return completed;
        }

        public bool CheckMarkedDrawingExists(string itemid)
        {
            bool itempresent = false;
            if (itemid == "")
            {
                itemid = Getfilename();
            }
            if (ValidfileName(itemid))
            {
                using (SqlCommand sqlCommand = new SqlCommand("SELECT COUNT(*) FROM [SPM_Database].[dbo].[DrawingApprovals] WHERE [ItemNo]='" + itemid.ToString() + "' AND  ISNULL([CheckedOn], '') <> '' ", cn))
                {
                    try
                    {
                        cn.Open();

                        int userCount = (int)sqlCommand.ExecuteScalar();
                        if (userCount == 1)
                        {
                            //MessageBox.Show("item already exists");
                            itempresent = true;
                        }
                        else
                        {
                            //MessageBox.Show(" move forward");
                            itempresent = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "SPM Connect - Check Item Present On Drawing Checked", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cn.Close();
                    }
                }
            }
            return itempresent;
        }

        private void MarkDrawingFile(ModelDoc2 swModel, bool approved)
        {
            try
            {
                string userid = Getuserfullname();
                DateTime datecreated = DateTime.Now;
                string sqlFormattedDate = datecreated.ToString("yyyy-MM-dd HH:mm tt");
                if (approved)
                {
                    swModel.DeleteCustomInfo("QAApproval");
                    swModel.DeleteCustomInfo("QAAppDate");
                    swModel.AddCustomInfo3("", "QAApproval", (int)swCustomInfoType_e.swCustomInfoText, userid);
                    swModel.AddCustomInfo3("", "QAAppDate", (int)swCustomInfoType_e.swCustomInfoText, sqlFormattedDate);
                }
                else
                {
                    swModel.DeleteCustomInfo("CheckedBy");
                    swModel.DeleteCustomInfo("CheckedDate");
                    swModel.DeleteCustomInfo("QAApproval");
                    swModel.DeleteCustomInfo("QAAppDate");
                    swModel.AddCustomInfo3("", "CheckedBy", (int)swCustomInfoType_e.swCustomInfoText, userid);
                    swModel.AddCustomInfo3("", "CheckedDate", (int)swCustomInfoType_e.swCustomInfoText, sqlFormattedDate);
                }

                swModel.ForceRebuild3(true);
                swModel.Save3(1, 0, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Mark On Drawings", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddDrawToChecked(string itemid)
        {
            string userid = Getuserfullname();
            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("dd-MM-yyyy HH:mm tt");
            userid += ",";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO [SPM_Database].[dbo].[DrawingApprovals] ([ItemNo],[CheckedOn],[CheckedBy]) VALUES('" + itemid + "','" + sqlFormattedDate + " ','" + userid + "')";
                cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Mark Drawing Checked SQL", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        private void UpdateDrawToApproved(string itemid)
        {
            string userid = Getuserfullname();
            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("dd-MM-yyyy HH:mm tt");
            userid += ",";
            try
            {
                if (cn.State == ConnectionState.Closed)
                    cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE [SPM_Database].[dbo].[DrawingApprovals] SET [ApprovedOn] = '" + sqlFormattedDate + "',[ApprovedBy] = '" + userid + "'  WHERE [ItemNo] = '" + itemid + "'";
                cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Mark Drawing Approved/ SQL", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        private void UpdateDrawingToSQLForChecked(string itemid)
        {
            string userid = Getuserfullname();
            DateTime datecreated = DateTime.Now;
            string sqlFormattedDate = datecreated.ToString("dd-MM-yyyy HH:mm tt");
            if (cn.State == ConnectionState.Closed)
                cn.Open();
            try
            {
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE [SPM_Database].[dbo].[DrawingApprovals] SET [CheckedOn] = '" + sqlFormattedDate + "',[CheckedBy] = '" + userid + "',[ApprovedOn] = Null,[ApprovedBy] = ''  WHERE [ItemNo] = '" + itemid + "'";
                cmd.ExecuteNonQuery();
                cn.Close();
                //MessageBox.Show("New entry created", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SPM Connect - Update Item on [DrawingApprovals]", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cn.Close();
            }
        }

        public void MarkDrawings(Form1 form1, bool approved)
        {
            int i = 1;
            ModelDoc2 swModel;
            ConfigurationManager swConfMgr;
            Configuration swConf;
            Component2 swRootComp;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                return;
            }
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                //grab job and subassy no
                string drawing = Checkforspmdrwfile(Getfilename());
                if (drawing != null)
                {
                    form1.UpdateProgressBar(50, "Currenty marking : " + drawing);
                    Open_drw(drawing);
                    ReloadSheetformat();
                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    MarkDrawingFile(swModel, approved);
                    if (approved)
                    {
                        AddToApprovedDrawing("");
                    }
                    else
                    {
                        AddToCheckedDrawing("");
                    }
                    swApp.QuitDoc(swModel.GetTitle());
                    form1.UpdateProgressBar(100, "Marking Complete." + (approved ? "Print can be sent for release and built." : " Print can be sent for approval now."));
                }
            }

            //TraverseModelFeatures(swModel, 1);
            List<BOM> bomlist = new List<BOM>();
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
                swConf = (Configuration)swConfMgr.ActiveConfiguration;
                swRootComp = (Component2)swConf.GetRootComponent();

                bomlist = TraverseComponent(swRootComp, 1);

                string assydrw = Checkforspmdrwfile(Getfilename());
                if (assydrw != null)
                {
                    form1.UpdateProgressBar(i, "Currenty marking : " + assydrw);

                    Open_drw(assydrw);
                    ReloadSheetformat();
                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    MarkDrawingFile(swModel, approved);
                    if (approved)
                    {
                        AddToApprovedDrawing("");
                    }
                    else
                    {
                        AddToCheckedDrawing("");
                    }
                    swApp.QuitDoc(swModel.GetTitle());
                }
                if (bomlist.Count > 0)
                {
                    List<string> drawings = new List<string>();

                    foreach (BOM bom in bomlist)
                    {
                        string drawing = Checkforspmdrwfile(bom.ItemNo);
                        if (drawing != null)
                        {
                            drawings.Add(drawing);
                        }
                    }

                    int incrementby = 100 / drawings.Count;
                    i += incrementby;
                    foreach (string draw in drawings)
                    {
                        form1.UpdateProgressBar(i, "Currenty marking : " + draw);
                        Open_drw(draw);
                        ReloadSheetformat();
                        swModel = (ModelDoc2)swApp.ActiveDoc;
                        MarkDrawingFile(swModel, approved);
                        if (approved)
                        {
                            AddToApprovedDrawing("");
                        }
                        else
                        {
                            AddToCheckedDrawing("");
                        }
                        swApp.QuitDoc(swModel.GetTitle());
                        i = i + incrementby;
                    }
                }

                form1.UpdateProgressBar(100, "Marking Complete." + (approved ? "Prints can be sent for release and built." : " Prints can be sent for approval now."));
            }
        }

        #endregion Checking and Approving Drawings
    }
}