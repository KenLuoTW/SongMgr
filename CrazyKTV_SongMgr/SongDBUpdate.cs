using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    public partial class MainForm : Form
    {
        private void CarzyktvDB_CheckDatabaseFile()
        {
            Global.CrazyktvDatabaseError = false;

            if (File.Exists(Global.CrazyktvDatabaseFile))
            {
                List<string> CrazyktvDBTableList = new List<string>(CommonFunc.GetOleDbTableList(Global.CrazyktvDatabaseFile, ""));
                foreach (string TableName in Global.CrazyktvDBTableList)
                {
                    if (CrazyktvDBTableList.IndexOf(TableName) < 0)
                    {
                        Global.CrazyktvDatabaseError = true;
                        Global.SongMgrDBVerErrorUIStatus = false;
                        break;
                    }
                }
                CrazyktvDBTableList = null;
            }
            else
            {
                Global.CrazyktvDatabaseError = true;
                Global.SongMgrDBVerErrorUIStatus = false;
            }
        }

        private void SongMgrDB_CheckDatabaseFile()
        {
            Global.SongMgrDatabaseError = false;

            if (File.Exists(Global.CrazyktvSongMgrDatabaseFile))
            {
                List<string> SongMgrDBTableList = new List<string>(CommonFunc.GetOleDbTableList(Global.CrazyktvSongMgrDatabaseFile, ""));
                foreach (string TableName in Global.SongMgrDBTableList)
                {
                    if (SongMgrDBTableList.IndexOf(TableName) < 0)
                    {
                        Global.SongMgrDatabaseError = true;
                        Global.SongMgrDBVerErrorUIStatus = false;
                        break;
                    }
                }
                SongMgrDBTableList = null;
            }
            else
                            {
                Global.SongMgrDatabaseError = true;
                Global.SongMgrDBVerErrorUIStatus = false;
            }
        }


        private void SongDBUpdate_CheckDatabaseFile()
        {
            Global.CrazyktvDatabaseStatus = false;
            Global.CrazyktvDatabaseError = false;
            Global.SongMgrDatabaseError = false;
            Global.CrazyktvDatabaseMaxDigitCode = true;

            CarzyktvDB_CheckDatabaseFile();
            SongMgrDB_CheckDatabaseFile();

            if (!Global.SongMgrDatabaseError)
            {
                string CashboxUpdDate = "";
                double SongDBVer = Convert.ToDouble(Global.CrazyktvSongDBVer);  
                string sqlStr = "select * from ktv_Version";
                using (DataTable dt = CommonFunc.GetOleDbDataTable(Global.CrazyktvSongMgrDatabaseFile, sqlStr, "")) {
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            CashboxUpdDate = row["CashboxUpdDate"].ToString();
                            SongDBVer = Convert.ToDouble(row["SongDB"].ToString());
                        }
                        Global.CashboxUpdDate = DateTime.Parse(CashboxUpdDate);
                        Global.CrazyktvSongDBVer = SongDBVer.ToString("F2");

                        this.BeginInvoke((Action)delegate ()
                        {
                            SongMaintenance_DBVer1Value_Label.Text = SongDBVer.ToString("F2") + " 版";
                            Cashbox_UpdDateValue_Label.Text = (CultureInfo.CurrentCulture.Name == "zh-TW") ? DateTime.Parse(CashboxUpdDate).ToLongDateString() : DateTime.Parse(CashboxUpdDate).ToShortDateString();
                        });
                    }
                }
            }

            if (!Global.CrazyktvDatabaseError)
            {
                List<string> collist = CommonFunc.GetDBColumnList(Global.CrazyktvDatabaseFile, "ktv_Song", null);
                foreach (string col in collist)
                {
                    Global.CrazyktvDatabaseStatus = (Global.ktvSongColumnsList.IndexOf(col) < 0) ? false : true;
                }

                foreach (string col in Global.ktvSongColumnsList)
                {
                    Global.CrazyktvDatabaseStatus = (collist.IndexOf(col) < 0) ? false : true;
                }

                if (Global.CrazyktvDatabaseStatus)
                {
                    SongDBUpdate_UpdateFinish();
                }
                else
                {
                    MainTabControl.SelectedIndex = MainTabControl.TabPages.IndexOf(SongMaintenance_TabPage);
                    SongMaintenance_TabControl.SelectedIndex = SongMaintenance_TabControl.TabPages.IndexOf(SongMaintenance_DBVer_TabPage);
                    SongMaintenance_DBVerTooltip_Label.Text = "偵測到資料庫結構更動,開始進行更新...";
                    var UpdateDBTask = Task.Factory.StartNew(() => SongDBUpdate_UpdateDatabaseFile());
                }
            }
            else
            {
                SongDBUpdate_UpdateFinish();

            }
        }



        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:必須檢閱 SQL 查詢中是否有安全性弱點")]
        private void SongDBUpdate_UpdateDatabaseFile()
        {
            Global.TimerStartTime = DateTime.Now;

            using (OleDbConnection conn = CommonFunc.OleDbOpenConn(Global.CrazyktvDatabaseFile, ""))
            {
                bool UpdateError = false;

                if (!Directory.Exists(Application.StartupPath + @"\SongMgr\Backup")) Directory.CreateDirectory(Application.StartupPath + @"\SongMgr\Backup");
                string SongDBBackupFile = SongDBBackupFile = Application.StartupPath + @"\SongMgr\Backup\" + DateTime.Now.ToLongDateString() + "_CrazySong.mdb";
                File.Copy(Global.CrazyktvDatabaseFile, SongDBBackupFile, true);

                List<string> CrazyktvDBTableList = new List<string>(CommonFunc.GetOleDbTableList(Global.CrazyktvDatabaseFile, ""));

                // 移除 ktv_AllSinger 資料表
                if (CrazyktvDBTableList.IndexOf("ktv_AllSinger") >= 0)
                {
                    using (OleDbCommand cmd = new OleDbCommand("drop table ktv_AllSinger", conn))
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                        catch
                        {
                            UpdateError = true;
                            this.BeginInvoke((Action)delegate ()
                            {
                                SongMaintenance_DBVerTooltip_Label.Text = "移除 ktv_AllSinger 資料表失敗,已還原為原本的資料庫檔案。";
                            });
                        }
                    }
                }

                // 移除 ktv_Version 資料表
                if (CrazyktvDBTableList.IndexOf("ktv_Version") >= 0)
                {
                    using (OleDbCommand cmd = new OleDbCommand("drop table ktv_Version", conn))
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                        catch
                        {
                            UpdateError = true;
                            this.BeginInvoke((Action)delegate ()
                            {
                                SongMaintenance_DBVerTooltip_Label.Text = "移除 ktv_Version 資料表失敗,已還原為原本的資料庫檔案。";
                            });
                        }
                    }
                }
                CrazyktvDBTableList.Clear();
                CrazyktvDBTableList = null;

                if (!UpdateError)
                {
                    bool UpdateKtvSong = false;
                    bool UpdateKtvSinger = false;
                    bool UpdatePhonetics = false;
                    bool UpdateLangauage = true;
                    bool UpdateFavorite = false;
                    bool AddSongReplayGainColumn = true;
                    bool AddSongCashboxIdColumn = true;
                    bool RemoveSongMeanVolumeColumn = false;
                    bool RemoveGodLiuColumn = false;
                    List<string> GodLiuColumnlist = new List<string>();

                    List<string> tablelist = new List<string>() { "ktv_Singer", "ktv_Phonetics", "ktv_Langauage", "ktv_User", "ktv_Favorite" };
                    String[] Restrictions = new String[4];
                    Restrictions[2] = "ktv_Song";
                    using (DataTable dt = conn.GetSchema("Columns", Restrictions))
                    {
                        foreach (string tablename in tablelist)
                        {
                            Restrictions[2] = tablename;
                            using (DataTable tb = conn.GetSchema("Columns", Restrictions))
                            {
                                foreach (DataRow row in tb.Rows)
                                {
                                    dt.ImportRow(row);
                                }
                            }
                        }
                        tablelist.Clear();
                        tablelist = null;

                        foreach (DataRow row in dt.Rows)
                        {
                            switch (row["COLUMN_NAME"].ToString())
                            {
                                case "Song_SongName":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "80") UpdateKtvSong = true;
                                    break;
                                case "Song_Singer":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "60") UpdateKtvSong = true;
                                    break;
                                case "Song_Spell":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "80") UpdateKtvSong = true;
                                    break;
                                case "Song_FileName":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "255") UpdateKtvSong = true;
                                    break;
                                case "Song_SpellNum":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "80") UpdateKtvSong = true;
                                    break;
                                case "Song_PenStyle":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "80") UpdateKtvSong = true;
                                    break;
                                case "Song_ReplayGain":
                                    AddSongReplayGainColumn = false;
                                    break;
                                case "Song_CashboxId":
                                    AddSongCashboxIdColumn = false;
                                    break;
                                case "Song_MeanVolume":
                                    RemoveSongMeanVolumeColumn = true;
                                    break;
                                case "Singer_Name":
                                case "Singer_Spell":
                                case "Singer_SpellNum":
                                case "Singer_PenStyle":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "60") UpdateKtvSinger = true;
                                    break;
                                case "PenStyle":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "40") UpdatePhonetics = true;
                                    break;
                                case "Langauage_KeyWord":
                                    UpdateLangauage = false;
                                    break;
                                case "User_Id":
                                    if (row["CHARACTER_MAXIMUM_LENGTH"].ToString() != "12") UpdateFavorite = true;
                                    break;
                                case "Song_SongNameFuzzy":
                                case "Song_SingerFuzzy":
                                case "Song_FuzzyVer":
                                case "DLspace":
                                case "Epasswd":
                                case "imgpath":
                                case "cashboxsongid":
                                case "cashboxdat":
                                case "holidaysongid":
                                    RemoveGodLiuColumn = true;
                                    GodLiuColumnlist.Add(row["COLUMN_NAME"].ToString());
                                    break;
                            }
                        }
                    }

                    string UpdateSqlStr = "";
                    if (UpdateKtvSong)
                    {
                        try
                        {
                            UpdateSqlStr = "select * into ktv_Song_temp from ktv_Song";
                            using (OleDbCommand cmd = new OleDbCommand(UpdateSqlStr, conn))
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }

                            UpdateSqlStr = "delete * from ktv_Song";
                            using (OleDbCommand cmd = new OleDbCommand(UpdateSqlStr, conn))
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }

                            List<string> cmdstrlist = new List<string>()
                            {
                                "alter table ktv_Song alter column Song_SongName TEXT(80) WITH COMPRESSION",
                                "alter table ktv_Song alter column Song_Singer TEXT(60) WITH COMPRESSION",
                                "alter table ktv_Song alter column Song_Spell TEXT(80) WITH COMPRESSION",
                                "alter table ktv_Song alter column Song_FileName TEXT(255) WITH COMPRESSION",
                                "alter table ktv_Song alter column Song_SpellNum TEXT(80) WITH COMPRESSION",
                                "alter table ktv_Song alter column Song_PenStyle TEXT(80) WITH COMPRESSION"
                            };

                            foreach (string cmdstr in cmdstrlist)
                            {
                                using (OleDbCommand cmd = new OleDbCommand(cmdstr, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                    cmd.Parameters.Clear();
                                }
                            }
                            cmdstrlist.Clear();
                            cmdstrlist = null;

                            UpdateSqlStr = "insert into ktv_Song select * from ktv_Song_temp";
                            using (OleDbCommand cmd = new OleDbCommand(UpdateSqlStr, conn))
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }

                            UpdateSqlStr = "drop table ktv_Song_temp";
                            using (OleDbCommand cmd = new OleDbCommand(UpdateSqlStr, conn))
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }
                        }
                        catch
                        {
                            UpdateError = true;
                            this.BeginInvoke((Action)delegate ()
                            {
                                SongMaintenance_DBVerTooltip_Label.Text = "更新歌曲資料表失敗,已還原為原本的資料庫檔案。";
                            });
                        }
                    }

                    if (UpdateKtvSinger)
                    {
                        List<string> cmdstrlist = new List<string>()
                        {
                            "alter table ktv_Singer alter column Singer_Name TEXT(60) WITH COMPRESSION",
                            "alter table ktv_Singer alter column Singer_Spell TEXT(60) WITH COMPRESSION",
                            "alter table ktv_Singer alter column Singer_SpellNum TEXT(60) WITH COMPRESSION",
                            "alter table ktv_Singer alter column Singer_PenStyle TEXT(60) WITH COMPRESSION"
                        };

                        try
                        {
                            foreach (string cmdstr in cmdstrlist)
                            {
                                using (OleDbCommand cmd = new OleDbCommand(cmdstr, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                    cmd.Parameters.Clear();
                                }
                            }
                            cmdstrlist.Clear();
                            cmdstrlist = null;
                        }
                        catch
                        {
                            UpdateError = true;
                            this.BeginInvoke((Action)delegate ()
                            {
                                SongMaintenance_DBVerTooltip_Label.Text = "更新歌手資料表失敗,已還原為原本的資料庫檔案。";
                            });

                        }
                    }

                    if (UpdatePhonetics)
                    {
                        using (OleDbCommand cmd = new OleDbCommand("alter table ktv_Phonetics alter column PenStyle TEXT(40) WITH COMPRESSION", conn))
                        {
                            try
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }
                            catch
                            {
                                UpdateError = true;
                                this.BeginInvoke((Action)delegate ()
                                {
                                    SongMaintenance_DBVerTooltip_Label.Text = "更新拼音資料表失敗,已還原為原本的資料庫檔案。";
                                });
                            }
                        }
                    }

                    if (UpdateLangauage)
                    {
                        using (OleDbCommand cmd = new OleDbCommand("alter table ktv_Langauage add column Langauage_KeyWord TEXT(255) WITH COMPRESSION", conn))
                        {
                            try
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }
                            catch
                            {
                                UpdateError = true;
                                this.BeginInvoke((Action)delegate ()
                                {
                                    SongMaintenance_DBVerTooltip_Label.Text = "更新語系資料表失敗,已還原為原本的資料庫檔案。";
                                });
                            }
                        }
                    }

                    if (UpdateFavorite)
                    {
                        List<string> cmdstrlist = new List<string>()
                        {
                            "alter table ktv_User alter column User_Id TEXT(12) WITH COMPRESSION",
                            "alter table ktv_Favorite alter column User_Id TEXT(12) WITH COMPRESSION",
                        };

                        try
                        {
                            foreach (string cmdstr in cmdstrlist)
                            {
                                using (OleDbCommand cmd = new OleDbCommand(cmdstr, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                    cmd.Parameters.Clear();
                                }
                            }
                            cmdstrlist.Clear();
                            cmdstrlist = null;
                        }
                        catch
                        {
                            UpdateError = true;
                            this.BeginInvoke((Action)delegate ()
                            {
                                SongMaintenance_DBVerTooltip_Label.Text = "更新我的最愛資料表失敗,已還原為原本的資料庫檔案。";
                            });

                        }
                    }

                    if (AddSongReplayGainColumn)
                    {
                        using (OleDbCommand cmd = new OleDbCommand("alter table ktv_Song add column Song_ReplayGain DOUBLE", conn))
                        {
                            try
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }
                            catch
                            {
                                UpdateError = true;
                                this.BeginInvoke((Action)delegate ()
                                {
                                    SongMaintenance_DBVerTooltip_Label.Text = "加入 Song_ReplayGain 欄位失敗,已還原為原本的資料庫檔案。";
                                });
                            }
                        }
                    }
                    //"alter table ktv_User alter column User_Id TEXT(12) WITH COMPRESSION",
                    if (AddSongCashboxIdColumn)
                    {
                        using (OleDbCommand cmd = new OleDbCommand("alter table ktv_Song add column Song_CashboxId TEXT(20) WITH COMPRESSION", conn))
                        {
                            try
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }
                            catch
                            {
                                UpdateError = true;
                                this.BeginInvoke((Action)delegate ()
                                {
                                    SongMaintenance_DBVerTooltip_Label.Text = "加入 Song_CashboxId 欄位失敗,已還原為原本的資料庫檔案。";
                                });
                            }
                        }
                    }

                    if (RemoveSongMeanVolumeColumn)
                    {
                        using (OleDbCommand cmd = new OleDbCommand("alter table ktv_Song drop column Song_MeanVolume", conn))
                        {
                            try
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();
                            }
                            catch
                            {
                                UpdateError = true;
                                this.BeginInvoke((Action)delegate ()
                                {
                                    SongMaintenance_DBVerTooltip_Label.Text = "移除 Song_MeanVolume 欄位失敗,已還原為原本的資料庫檔案。";
                                });
                            }
                        }
                    }

                    if (RemoveGodLiuColumn)
                    {
                        List<string> haveindexlist = new List<string>();
                        using (DataTable dt = conn.GetSchema("Indexes"))
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                if (haveindexlist.IndexOf(row["COLUMN_NAME"].ToString()) < 0) haveindexlist.Add(row["COLUMN_NAME"].ToString());
                            }
                        }

                        string cmdstr = string.Empty;
                        string removeindex = string.Empty;
                        foreach (string GodLiuColumn in GodLiuColumnlist)
                        {
                            switch (GodLiuColumn)
                            {
                                case "Song_SongNameFuzzy":
                                    cmdstr = "alter table ktv_Song drop column Song_SongNameFuzzy";
                                    break;
                                case "Song_SingerFuzzy":
                                    cmdstr = "alter table ktv_Song drop column Song_SingerFuzzy";
                                    break;
                                case "Song_FuzzyVer":
                                    cmdstr = "alter table ktv_Song drop column Song_FuzzyVer";
                                    break;
                                case "DLspace":
                                    cmdstr = "alter table ktv_Song drop column DLspace";
                                    break;
                                case "Epasswd":
                                    cmdstr = "alter table ktv_Song drop column Epasswd";
                                    break;
                                case "imgpath":
                                    cmdstr = "alter table ktv_Song drop column imgpath";
                                    break;
                                case "cashboxsongid":
                                    removeindex = "drop index cashboxsongid on ktv_Song";
                                    cmdstr = "alter table ktv_Song drop column cashboxsongid";
                                    break;
                                case "cashboxdat":
                                    cmdstr = "alter table ktv_Song drop column cashboxdat";
                                    break;
                                case "holidaysongid":
                                    removeindex = "drop index holidaysongid on ktv_Song";
                                    cmdstr = "alter table ktv_Song drop column holidaysongid";
                                    break;
                            }

                            if (GodLiuColumn == "cashboxsongid" || GodLiuColumn == "holidaysongid")
                            {
                                if (haveindexlist.IndexOf(GodLiuColumn) > 0)
                                {
                                    using (OleDbCommand cmd = new OleDbCommand(removeindex, conn))
                                    {
                                        try
                                        {
                                            cmd.ExecuteNonQuery();
                                            cmd.Parameters.Clear();
                                        }
                                        catch
                                        {
                                            UpdateError = true;
                                            this.BeginInvoke((Action)delegate ()
                                            {
                                                SongMaintenance_DBVerTooltip_Label.Text = "刪除 GodLiu 相關欄位失敗,已還原為原本的資料庫檔案。";
                                            });
                                        }
                                    }
                                }
                            }

                            using (OleDbCommand cmd = new OleDbCommand(cmdstr, conn))
                            {
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                    cmd.Parameters.Clear();
                                }
                                catch
                                {
                                    UpdateError = true;
                                    this.BeginInvoke((Action)delegate ()
                                    {
                                        SongMaintenance_DBVerTooltip_Label.Text = "刪除 GodLiu 相關欄位失敗,已還原為原本的資料庫檔案。";
                                    });
                                }
                            }
                        }
                    }

                    string SqlStr = "select * from ktv_Swan";
                    using (DataTable dt = CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, SqlStr, ""))
                    {
                        if (dt.Rows.Count > 0)
                        {
                            if (Convert.ToString(dt.Rows[3][1]) == "合唱歌曲")
                            {
                                using (OleDbCommand cmd = new OleDbCommand("update ktv_Swan set Swan_Name = @SwanName where Swan_Id = @SwanId", conn))
                                {
                                    cmd.Parameters.AddWithValue("@SwanName", "合唱");
                                    cmd.Parameters.AddWithValue("@SwanId", "3");
                                    try
                                    {
                                        cmd.ExecuteNonQuery();
                                        cmd.Parameters.Clear();
                                    }
                                    catch
                                    {
                                        UpdateError = true;
                                        this.BeginInvoke((Action)delegate ()
                                        {
                                            SongMaintenance_DBVerTooltip_Label.Text = "變更歌手類別資料失敗,已還原為原本的資料庫檔案。";
                                        });
                                    }
                                }
                            }
                        }
                    }
                }

                if (UpdateError)
                {
                    File.Copy(SongDBBackupFile, Global.CrazyktvDatabaseFile, true);
                    Global.DatabaseUpdateFinished = true;
                }
                else
                {
                    CommonFunc.SaveConfigXmlFile(Global.SongMgrCfgFile, "CrazyktvSongDBVer", Global.CrazyktvSongDBVer);

                    string VersionQuerySqlStr = "select * from ktv_Version";
                    using (DataTable dt = CommonFunc.GetOleDbDataTable(Global.CrazyktvSongMgrDatabaseFile, VersionQuerySqlStr, ""))
                    {
                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                Global.CashboxUpdDate = DateTime.Parse(row["CashboxUpdDate"].ToString());
                            }
                        }
                    }

                    this.BeginInvoke((Action)delegate ()
                    {
                        Global.TimerEndTime = DateTime.Now;

                        SongMaintenance_DBVer1Value_Label.Text = Global.CrazyktvSongDBVer + " 版";
                        Cashbox_UpdDateValue_Label.Text = (CultureInfo.CurrentCulture.Name == "zh-TW") ? Global.CashboxUpdDate.ToLongDateString() : Global.CashboxUpdDate.ToShortDateString();

                        SongMaintenance_DBVerTooltip_Label.Text = "";
                        SongMaintenance_Tooltip_Label.Text = "已完成歌庫版本更新,共花費 " + (long)(Global.TimerEndTime - Global.TimerStartTime).TotalSeconds + " 秒完成。";
                    });
                    Global.CrazyktvDatabaseStatus = true;
                    SongDBUpdate_UpdateFinish();
                }
            }
        }


        private void SongDBUpdate_UpdateFinish()
        {
            if (Global.CrazyktvDatabaseStatus)
            {
                DataTable dt = new DataTable();
                string SongQuerySqlStr = "select Song_Id from ktv_Song";
                dt = CommonFunc.GetOleDbDataTable(Global.CrazyktvDatabaseFile, SongQuerySqlStr, "");
                if (dt.Rows.Count > 0)
                {
                    var d5code = from row in dt.AsEnumerable()
                                 where row.Field<string>("Song_Id").Length == 5
                                 select row;

                    var d6code = from row in dt.AsEnumerable()
                                 where row.Field<string>("Song_Id").Length == 6
                                 select row;

                    int MaxDigitCode;
                    if (d5code.Count<DataRow>() > d6code.Count<DataRow>()) { MaxDigitCode = 5; } else { MaxDigitCode = 6; }

                    switch (MaxDigitCode)
                    {
                        case 5:
                            ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb => cb.Enabled = false);
                            if (Global.SongMgrMaxDigitCode != "1")
                            {
                                ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb => cb.SelectedValue = 1);
                                CommonFunc.SaveConfigXmlFile(Global.SongMgrCfgFile, "SongMgrMaxDigitCode", Global.SongMgrMaxDigitCode);
                                CommonFunc.SaveConfigXmlFile(Global.SongMgrCfgFile, "SongMgrLangCode", Global.SongMgrLangCode);
                            }
                            break;
                        case 6:
                            ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb => cb.Enabled = false);
                            if (Global.SongMgrMaxDigitCode != "2")
                            {
                                ControlExtensions.BeginInvokeIfRequired(SongMgrCfg_MaxDigitCode_ComboBox, cb => cb.SelectedValue = 2);
                                CommonFunc.SaveConfigXmlFile(Global.SongMgrCfgFile, "SongMgrMaxDigitCode", Global.SongMgrMaxDigitCode);
                                CommonFunc.SaveConfigXmlFile(Global.SongMgrCfgFile, "SongMgrLangCode", Global.SongMgrLangCode);
                            }
                            break;
                    }

                    var query = from row in dt.AsEnumerable()
                                where row.Field<string>("Song_Id").Length != MaxDigitCode
                                select row;

                    if (query.Count<DataRow>() > 0)
                    {
                        Global.SongMgrDBVerErrorUIStatus = false;
                        ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = false);
                        ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeCorrect_Button, btn => btn.Enabled = true);
                        Global.CrazyktvDatabaseMaxDigitCode = false;
                        Global.CrazyktvDatabaseStatus = false;
                    }
                    else
                    {
                        if (Global.SongMgrSongAddMode == "3" || Global.SongMgrSongAddMode == "4")
                        {
                            Global.SongMgrDBVerErrorUIStatus = true;
                        }
                        else
                        {
                            if (Directory.Exists(Global.SongMgrDestFolder))
                            {
                                Global.SongMgrDBVerErrorUIStatus = true;
                            } else
                            {
                                Global.SongMgrDBVerErrorUIStatus = false;
                                Global.CrazyktvDatabaseStatus = false;
                            }
                        }

                        switch (Global.SongMgrMaxDigitCode)
                        {
                            case "1":
                                ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = true);
                                break;
                            case "2":
                                ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = false);
                                break;
                        }
                        ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeCorrect_Button, btn => btn.Enabled = false);
                        Global.CrazyktvDatabaseMaxDigitCode = true;
                    }
                }
                else // 空白資料庫
                {
                    if (Global.SongMgrSongAddMode == "3" || Global.SongMgrSongAddMode == "4")
                    {
                        Global.SongMgrDBVerErrorUIStatus = true;
                    }
                    else
                    {
                        if (Directory.Exists(Global.SongMgrDestFolder)) { Global.SongMgrDBVerErrorUIStatus = true; } else { Global.SongMgrDBVerErrorUIStatus = false; }
                    }

                    ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeConvTo6_Button, btn => btn.Enabled = false);
                    ControlExtensions.BeginInvokeIfRequired(SongMaintenance_CodeCorrect_Button, btn => btn.Enabled = false);
                    Global.CrazyktvDatabaseMaxDigitCode = true;
                }
                dt.Dispose();
                dt = null;
            }

            Console.WriteLine();

            if (Global.CrazyktvDatabaseStatus)
            {
                // 檢查是否有自訂語系
                Common_CheckSongLang();

                // 統計歌曲數量
                Task.Factory.StartNew(() => Common_GetSongStatisticsTask());

                // 統計歌手數量
                Task.Factory.StartNew(() => Common_GetSingerStatisticsTask());

                // 檢查備份移除歌曲是否要刪除
                Task.Factory.StartNew(() => Common_CheckBackupRemoveSongTask());

                // 取得可用歌曲編號
                Task.Factory.StartNew(() => CommonFunc.GetRemainingSongIdCount((Global.SongMgrMaxDigitCode == "1") ? 5 : 6));

                this.BeginInvoke((Action)delegate()
                {
                    // 載入我的最愛清單
                    SongQuery_GetFavoriteUserList();

                    // 歌庫設定 - 載入下拉選單清單及設定
                    SongMgrCfg_SetLangLB();

                    // 歌庫維護 - 載入下拉選單清單及設定
                    SongMaintenance_GetFavoriteUserList();
                    SongMaintenance_SetCustomLangControl();
                });
            }
            Global.DatabaseUpdateFinished = true;
        }

    }
}
