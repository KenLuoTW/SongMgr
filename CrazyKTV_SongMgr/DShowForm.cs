﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading.Tasks;
using Vlc.DotNet.Forms;
using System.IO;
using Vlc.DotNet.Core;
using System.Windows.Forms.Integration;
using CrazyKTV_MediaKit.DirectShow.Controls;
using System.Windows;
using System.Threading;

namespace CrazyKTV_SongMgr
{
    public partial class DShowForm : Form
    {
        string SongId;
        string SongLang;
        string SongSinger;
        string SongSongName;
        string SongTrack;
        string SongVolume;
        string SongReplayGain;
        string SongMeanVolume;
        string SongFilePath;
        string dvRowIndex;
        string UpdateSongTrack;
        string UpdateDataGridView;

        private MediaUriElement mediaUriElement;
        private bool sliderInit;
        private bool sliderDrag;

        public DShowForm()
        {
            InitializeComponent();
        }

        public DShowForm(Form ParentForm, List<string> PlayerSongInfoList)
        {
            InitializeComponent();

            this.Owner = ParentForm;
            SongId = PlayerSongInfoList[0];
            SongLang = PlayerSongInfoList[1];
            SongSinger = PlayerSongInfoList[2];
            SongSongName = PlayerSongInfoList[3];
            SongTrack = PlayerSongInfoList[4];
            SongVolume = PlayerSongInfoList[5];
            SongReplayGain = PlayerSongInfoList[6];
            SongMeanVolume = PlayerSongInfoList[7];
            SongFilePath = PlayerSongInfoList[8];
            dvRowIndex = PlayerSongInfoList[9];
            UpdateDataGridView = PlayerSongInfoList[10];

            this.Text = "【" + SongLang + "】" + SongSinger + " - " + SongSongName;

            sliderInit = false;

            mediaUriElement = new MediaUriElement();
            mediaUriElement.BeginInit();
            elementHost.Child = mediaUriElement;
            mediaUriElement.EndInit();

            mediaUriElement.MediaUriPlayer.CodecsDirectory = System.Windows.Forms.Application.StartupPath + @"\Codec";
            mediaUriElement.VideoRenderer = (Global.MainCfgPlayerOutput == "1") ? CrazyKTV_MediaKit.DirectShow.MediaPlayers.VideoRendererType.VideoMixingRenderer9 : CrazyKTV_MediaKit.DirectShow.MediaPlayers.VideoRendererType.EnhancedVideoRenderer;
            mediaUriElement.DeeperColor = (Global.MainCfgPlayerOutput == "1") ? false : true;
            mediaUriElement.Stretch = System.Windows.Media.Stretch.Fill;
            mediaUriElement.EnableAudioCompressor = bool.Parse(Global.MainCfgPlayerEnableAudioCompressor);
            mediaUriElement.EnableAudioProcessor = true;

            mediaUriElement.MediaFailed += MediaUriElement_MediaFailed;
            mediaUriElement.MediaEnded += MediaUriElement_MediaEnded;
            mediaUriElement.MouseLeftButtonDown += mediaUriElement_MouseLeftButtonDown;
            mediaUriElement.MediaUriPlayer.MediaPositionChanged += MediaUriPlayer_MediaPositionChanged;

            mediaUriElement.Source = new Uri(SongFilePath);

            // 音量平衡
            int GainVolume = Convert.ToInt32(SongVolume);
            if (SongReplayGain != "" && SongMeanVolume != "")
            {
                int basevolume = 100;
                GainVolume = basevolume;

                List<int> maxvolumelist = new List<int>() { -18, -17, -16, -15, -14, -13, -12, -11, -10, -9 };
                int maxvolume = maxvolumelist[Convert.ToInt32(Global.SongMaintenanceMaxVolume) - 1];
                maxvolumelist.Clear();
                maxvolumelist = null;

                double GainDB = Convert.ToDouble(SongReplayGain);
                double MeanDB = Convert.ToDouble(SongMeanVolume);
                if (GainDB * -1 > 0)
                {
                    GainVolume = Convert.ToInt32(basevolume * Math.Pow(10, (GainDB * -1) / 20));
                }
                else
                {
                    if (MeanDB > maxvolume)
                    {
                        GainVolume = Convert.ToInt32(basevolume * Math.Pow(10, (maxvolume - MeanDB) / 20));
                    }
                }
            }
            mediaUriElement.AudioAmplify = GainVolume;
            Player_CurrentGainValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = GainVolume + " %");

            SpinWait.SpinUntil(() => mediaUriElement.GetAudioTrackList().Count > 0);

            mediaUriElement.AudioTrackList = mediaUriElement.GetAudioTrackList();
            string ChannelValue = string.Empty;
            if (mediaUriElement.AudioTrackList.Count == 1)
            {
                switch (SongTrack)
                {
                    case "1":
                        if (mediaUriElement.AudioChannel != 1) mediaUriElement.AudioChannel = 1;
                        ChannelValue = "1";
                        break;
                    case "2":
                        if (mediaUriElement.AudioChannel != 2) mediaUriElement.AudioChannel = 2;
                        ChannelValue = "2";
                        break;
                }
            }
            else if (mediaUriElement.AudioTrackList.Count > 1)
            {
                switch (SongTrack)
                {
                    case "1":
                        if (mediaUriElement.AudioTrackList.IndexOf(mediaUriElement.AudioTrack) != mediaUriElement.AudioTrackList[0]) mediaUriElement.AudioTrack = mediaUriElement.AudioTrackList[0];
                        ChannelValue = "1";
                        break;
                    case "2":
                        if (mediaUriElement.AudioTrackList.IndexOf(mediaUriElement.AudioTrack) != mediaUriElement.AudioTrackList[1]) mediaUriElement.AudioTrack = mediaUriElement.AudioTrackList[1];
                        ChannelValue = "2";
                        break;
                }
            }
            Player_CurrentChannelValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = (ChannelValue == SongTrack) ? "伴唱" : "人聲");

            NativeMethods.SystemSleepManagement.PreventSleep(true);
        }

        private void MediaUriElement_MediaFailed(object sender, CrazyKTV_MediaKit.DirectShow.MediaPlayers.MediaFailedEventArgs e)
        {
            this.BeginInvokeIfRequired(form => form.Text = e.Message);
        }

        private void MediaUriPlayer_MediaPositionChanged(object sender, EventArgs e)
        {
            if (sliderDrag)
                return;

            if (!sliderInit)
            {
                this.Invoke((Action)delegate ()
                {
                    if (mediaUriElement.MediaDuration > 0)
                    {
                        Player_ProgressTrackBar.Maximum = ((int)mediaUriElement.MediaDuration < 0) ? (int)mediaUriElement.MediaDuration * -1 : (int)mediaUriElement.MediaDuration;
                        sliderInit = true;
                    }
                });
            }
            else
            {
                this.BeginInvoke(new Action(ChangeSlideValue), null);
            }
        }

        private void ChangeSlideValue()
        {
            if (sliderDrag)
                return;

            if (sliderInit)
            {
                double perc = (double)mediaUriElement.MediaPosition / mediaUriElement.MediaDuration;
                int newValue = (int)(Player_ProgressTrackBar.Maximum * perc);
                if (newValue - Player_ProgressTrackBar.ProgressBarValue < 500000) return;
                Player_ProgressTrackBar.TrackBarValue = newValue;
                Player_ProgressTrackBar.ProgressBarValue = newValue;
            }
        }

        private void Player_ProgressTrackBar_Click(object sender, EventArgs e)
        {
            if (!sliderInit)
                return;

            this.BeginInvoke(new Action(ChangeMediaPosition), null);
        }

        private void ChangeMediaPosition()
        {
            sliderDrag = true;
            double perc = (double)Player_ProgressTrackBar.TrackBarValue / Player_ProgressTrackBar.Maximum;
            mediaUriElement.MediaPosition = (long)(mediaUriElement.MediaDuration * perc);
            Player_ProgressTrackBar.ProgressBarValue = Player_ProgressTrackBar.TrackBarValue;
            sliderDrag = false;
        }

        private void MediaUriElement_MediaEnded(object sender, RoutedEventArgs e)
        {
            mediaUriElement.Stop();
            mediaUriElement.MediaPosition = 0;
            Player_ProgressTrackBar.TrackBarValue = 0;
            Player_ProgressTrackBar.ProgressBarValue = 0;
            Player_PlayControl_Button.Text = "播放";
        }

        private void Player_SwithChannel_Button_Click(object sender, EventArgs e)
        {
            string ChannelValue = string.Empty;

            if (mediaUriElement.AudioTrackList.Count > 1)
            {
                if (mediaUriElement.AudioTrackList.IndexOf(mediaUriElement.AudioTrack) == 0)
                {
                    mediaUriElement.AudioTrack = mediaUriElement.AudioTrackList[1];
                    ChannelValue = "2";
                    UpdateSongTrack = "2";
                }
                else
                {
                    mediaUriElement.AudioTrack = mediaUriElement.AudioTrackList[0];
                    ChannelValue = "1";
                    UpdateSongTrack = "1";
                }
            }
            else
            {
                if (mediaUriElement.AudioChannel == 1)
                {
                    mediaUriElement.AudioChannel = 2;
                    ChannelValue = "2";
                    UpdateSongTrack = "2";
                }
                else
                {
                    mediaUriElement.AudioChannel = 1;
                    ChannelValue = "1";
                    UpdateSongTrack = "1";
                }
            }
            Player_CurrentChannelValue_Label.Text = (ChannelValue == SongTrack) ? "伴唱" : "人聲";
            Player_UpdateChannel_Button.Enabled = (Player_CurrentChannelValue_Label.Text == "人聲") ? true : false;
        }

        private void Player_UpdateChannel_Button_Click(object sender, EventArgs e)
        {
            SongTrack = UpdateSongTrack;
            Player_UpdateChannel_Button.Enabled = false;
            Player_CurrentChannelValue_Label.Text = "伴唱";
            Global.PlayerUpdateSongValueList = new List<string>() { UpdateDataGridView, dvRowIndex, SongTrack };
        }

        private void Player_PlayControl_Button_Click(object sender, EventArgs e)
        {
            switch (((Button)sender).Text)
            {
                case "暫停播放":
                    mediaUriElement.Pause();
                    ((Button)sender).Text = "繼續播放";
                    break;
                case "繼續播放":
                    mediaUriElement.Play();
                    ((Button)sender).Text = "暫停播放";
                    break;
                case "播放":
                    mediaUriElement.Play();
                    ((Button)sender).Text = "暫停播放";
                    break;
            }
        }

        private void mediaUriElement_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                ToggleFullscreen();
            }
        }

        private FormWindowState winState;
        private System.Drawing.Point winLoc;
        private int winWidth;
        private int winHeight;
        private int eHostWidth;
        private int eHostHeight;
        
        private void ToggleFullscreen()
        {
            if (this.FormBorderStyle == FormBorderStyle.None)
            {
                this.WindowState = winState;

                this.Hide();
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.TopMost = false;
                this.Location = winLoc;
                this.Width = winWidth;
                this.Height = winHeight;

                elementHost.Dock = DockStyle.None;
                elementHost.Location = new System.Drawing.Point(12, 12);
                elementHost.Width = eHostWidth;
                elementHost.Height = eHostHeight;
                elementHost.Anchor = AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Left;
                this.Show();
            }
            else
            {
                winState = this.WindowState;
                winLoc = this.Location;
                winWidth = this.Width;
                winHeight = this.Height;
                eHostWidth = elementHost.Width;
                eHostHeight = elementHost.Height;

                this.Hide();
                this.FormBorderStyle = FormBorderStyle.None;
                this.TopMost = true;
                this.WindowState = FormWindowState.Normal;
                this.WindowState = FormWindowState.Maximized;
                elementHost.Dock = DockStyle.Fill;
                this.Show();
            }
        }

        private void DShowForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            mediaUriElement.Stop();
            mediaUriElement.Close();
            mediaUriElement.Source = null;
            
            NativeMethods.SystemSleepManagement.ResotreSleep();
            this.Owner.Show();
        }

        private void DShowForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
            GC.Collect();
        }
    }
}