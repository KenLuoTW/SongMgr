﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms;
using System.Linq;
using System.Threading;
using CrazyKTV_MediaKit.DirectShow.Controls;
using CrazyKTV_MediaKit.DirectShow.MediaPlayers;

namespace CrazyKTV_SongMgr
{
    public partial class DShowForm : Form
    {
        private string SongId;
        private string SongLang;
        private string SongSinger;
        private string SongSongName;
        private string SongTrack;
        private string SongVolume;
        private string SongReplayGain;
        private string SongFilePath;
        private string dvRowIndex;
        private string UpdateSongTrack;
        private string UpdateDataGridView;

        private MediaUriElement mediaUriElement;
        private System.Timers.Timer mouseClickTimer;
        private DateTime MediaPositionChangeTime;
        private bool sliderInit;
        private bool sliderDrag;

        public DShowForm()
        {
            InitializeComponent();
        }

        public DShowForm(Form ParentForm, List<string> PlayerSongInfoList)
        {
            InitializeComponent();
            this.MouseWheel += new MouseEventHandler(DShowForm_MouseWheel);

            this.Owner = ParentForm;
            SongId = PlayerSongInfoList[0];
            SongLang = PlayerSongInfoList[1];
            SongSinger = PlayerSongInfoList[2];
            SongSongName = PlayerSongInfoList[3];
            SongTrack = PlayerSongInfoList[4];
            SongVolume = PlayerSongInfoList[5];
            SongReplayGain = PlayerSongInfoList[6];
            SongFilePath = PlayerSongInfoList[7];
            dvRowIndex = PlayerSongInfoList[8];
            UpdateDataGridView = PlayerSongInfoList[9];

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
            mediaUriElement.EnableAudioProcessor = false;

            mediaUriElement.MediaOpened += MediaUriElement_MediaOpened;
            mediaUriElement.MediaFailed += MediaUriElement_MediaFailed;
            mediaUriElement.MediaEnded += MediaUriElement_MediaEnded;
            mediaUriElement.MouseLeftButtonDown += mediaUriElement_MouseLeftButtonDown;
            mediaUriElement.MediaUriPlayer.MediaPositionChanged += MediaUriPlayer_MediaPositionChanged;
            MediaPositionChangeTime = DateTime.Now;

            // 隨選視訊
            if (Global.PlayerRandomVideoList.Count == 0)
            {
                string dir = System.Windows.Forms.Application.StartupPath + @"\Video";
                if (Directory.Exists(dir))
                {
                    Global.PlayerRandomVideoList.AddRange(Directory.GetFiles(dir));
                    if (Global.PlayerRandomVideoList.Count > 0)
                    {
                        Random rand = new Random(Guid.NewGuid().GetHashCode());
                        Global.PlayerRandomVideoList = Global.PlayerRandomVideoList.OrderBy(str => rand.Next()).ToList<string>();
                    }
                }
            }
            mediaUriElement.VideoSource = (Global.PlayerRandomVideoList.Count > 0) ? new Uri(Global.PlayerRandomVideoList[0]) : null;
            mediaUriElement.Source = new Uri(SongFilePath);

            mediaUriElement.Volume = Math.Round(Convert.ToDouble(Global.MainCfgPlayerDefaultVolume) / 100, 2);
            // 音量平衡
            int GainVolume = Convert.ToInt32(SongVolume);
            if (!string.IsNullOrEmpty(SongReplayGain))
            {
                int basevolume = 100;
                GainVolume = basevolume;

                double GainDB = Convert.ToDouble(SongReplayGain);
                GainVolume = Convert.ToInt32(basevolume * Math.Pow(10, GainDB / 20));

            }
            mediaUriElement.AudioAmplify = GainVolume;
            Player_CurrentGainValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = GainVolume + " %");

            if (mediaUriElement.MediaUriPlayer.IsAudioOnly && Global.PlayerRandomVideoList.Count > 0)
                Global.PlayerRandomVideoList.RemoveAt(0);

            NativeMethods.SystemSleepManagement.PreventSleep(true);
        }

        private void DShowForm_MouseWheel(object sender, MouseEventArgs e)
        {
            if (e.Delta != 0)
            {
                if (e.Delta > 0)
                {
                    if (mediaUriElement.Volume <= 0.95)
                    {
                        mediaUriElement.Volume += 0.05;
                    }
                    else
                    {
                        mediaUriElement.Volume = 1.00;
                    }
                }
                else
                {
                    if (mediaUriElement.Volume >= 0.05)
                    {
                        mediaUriElement.Volume -= 0.05;
                    }
                    else
                    {
                        mediaUriElement.Volume = 0;
                    }
                }
                mediaUriElement.Volume = Math.Round(mediaUriElement.Volume, 2);
                Player_CurrentVolumeValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = Convert.ToInt32(mediaUriElement.Volume * 100).ToString());
            }
        }

        private void MediaUriElement_MediaOpened(object sender, RoutedEventArgs e)
        {
            string ChannelValue = string.Empty;
            if (mediaUriElement.AudioStreams.Count == 1)
            {
                mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
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
            else if (mediaUriElement.AudioStreams.Count > 1)
            {
                switch (SongTrack)
                {
                    case "1":
                        if (Global.SongMgrSongTrackMode == "True")
                        {
                            mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
                        }
                        else
                        {
                            mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[1];
                        }
                        ChannelValue = "1";
                        break;
                    case "2":
                        if (Global.SongMgrSongTrackMode == "True")
                        {
                            mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[1];
                        }
                        else
                        {
                            mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
                        }
                        ChannelValue = "2";
                        break;
                }
            }
            Player_CurrentChannelValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = (ChannelValue == SongTrack) ? "伴唱" : "人聲");
            Player_CurrentVolumeValue_Label.BeginInvokeIfRequired(lbl => lbl.Text = Convert.ToInt32(mediaUriElement.Volume * 100).ToString());
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
                    if (mediaUriElement.HasVideo)
                    {
                        if (mediaUriElement.NaturalVideoWidth != 0 && mediaUriElement.NaturalVideoHeight != 0)
                        {
                            Player_VideoSizeValue_Label.Text = mediaUriElement.NaturalVideoWidth + "x" + mediaUriElement.NaturalVideoHeight;
                        }
                    }

                    if (mediaUriElement.MediaDuration > 0)
                    {
                        Player_ProgressTrackBar.Maximum = ((int)mediaUriElement.MediaDuration < 0) ? (int)mediaUriElement.MediaDuration * -1 : (int)mediaUriElement.MediaDuration;
                        sliderInit = true;
                    }
                });
            }
            else
            {
                if ((DateTime.Now - MediaPositionChangeTime).TotalMilliseconds < 500) return;
                this.BeginInvoke(new Action(ChangeSlideValue), null);
                MediaPositionChangeTime = DateTime.Now;
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

            ChangeMediaPosition();
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
            string ChannelValue;

            if (mediaUriElement.AudioStreams.Count > 1)
            {
                if (mediaUriElement.AudioTrack == mediaUriElement.AudioStreams[0])
                {
                    mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[1];
                    ChannelValue = (Global.SongMgrSongTrackMode == "True") ? "2" : "1";
                    UpdateSongTrack = (Global.SongMgrSongTrackMode == "True") ? "2" : "1";
                }
                else
                {
                    mediaUriElement.AudioTrack = mediaUriElement.AudioStreams[0];
                    ChannelValue = (Global.SongMgrSongTrackMode == "True") ? "1" : "2";
                    UpdateSongTrack = (Global.SongMgrSongTrackMode == "True") ? "1" : "2";
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
            if (mouseClickTimer == null)
            {
                mouseClickTimer = new System.Timers.Timer
                {
                    Interval = SystemInformation.DoubleClickTime
                };
                mouseClickTimer.Elapsed += new System.Timers.ElapsedEventHandler(mouseClickTimer_Tick);
            }

            if (!mouseClickTimer.Enabled)
            {
                mouseClickTimer.Start();
            }
            else
            {
                mouseClickTimer.Stop();
                ToggleFullscreen();
            }
        }

        private void mouseClickTimer_Tick(object sender, EventArgs e)
        {
            mouseClickTimer.Stop();

            switch (Player_PlayControl_Button.Text)
            {
                case "暫停播放":
                    mediaUriElement.Dispatcher.Invoke(new Action(() => mediaUriElement.Pause()));
                    Player_PlayControl_Button.InvokeIfRequired<Button>(btn => btn.Text = "繼續播放");
                    break;
                case "繼續播放":
                    mediaUriElement.Dispatcher.Invoke(new Action(() => mediaUriElement.Play()));
                    Player_PlayControl_Button.InvokeIfRequired<Button>(btn => btn.Text = "暫停播放");
                    break;
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
                Player_ProgressTrackBar.TrackBarValue = Player_ProgressTrackBar.TrackBarValue;
                Player_ProgressTrackBar.ProgressBarValue = Player_ProgressTrackBar.ProgressBarValue;
                Cursor.Show();
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
                Cursor.Hide();
                this.Show();
            }
        }

        private void DShowForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            mediaUriElement.MediaUriPlayer.MediaPositionChanged -= MediaUriPlayer_MediaPositionChanged;
            mediaUriElement.Stop();
            mediaUriElement.Close();
            mediaUriElement.Source = null;
            mediaUriElement.VideoSource = null;

            if (mouseClickTimer != null)
                mouseClickTimer.Dispose();

            NativeMethods.SystemSleepManagement.ResotreSleep();
            this.Owner.Show();
            this.Owner.TopMost = (Global.MainCfgAlwaysOnTop == "True") ? true : false;
        }

        private void DShowForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            mediaUriElement = null;
            GC.Collect();
        }
    }
}
