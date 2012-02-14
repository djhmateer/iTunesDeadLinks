using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using System.Collections;
using iTunesLib;

namespace iTunesCOMSample
{
    public partial class Form1 : Form
    {
        private volatile bool _shouldStop;
        private Thread worker;

        public Form1()
        {
            InitializeComponent();
        }

        private void RemoveDuplicates()
        {
            var iTunes = new iTunesAppClass();

            //get a reference to the collection of all tracks
            IITTrackCollection tracks = iTunes.LibraryPlaylist.Tracks;

            int trackCount = tracks.Count;
            int numberChecked = 0;
            int numberDuplicateFound = 0;
            Dictionary<string, IITTrack> trackCollection = new Dictionary<string, IITTrack>();
            var tracksToRemove = new ArrayList();

            //setup the progress control
            SetupProgress(trackCount);

            for (int i = trackCount; i > 0; i--)
            {
                if (tracks[i].Kind == ITTrackKind.ITTrackKindFile)
                {
                    if (!_shouldStop)
                    {
                        numberChecked++;
                        IncrementProgress();
                        UpdateLabel("Checking track # " + numberChecked.ToString() + " - " + tracks[i].Name);
                        string trackKey = tracks[i].Name + tracks[i].Artist + tracks[i].Album;

                        if (!trackCollection.ContainsKey(trackKey))
                            trackCollection.Add(trackKey, tracks[i]);
                        else
                        {
                            if (trackCollection[trackKey].Album != tracks[i].Album || trackCollection[trackKey].Artist != tracks[i].Artist)
                                trackCollection.Add(trackKey, tracks[i]);
                            else if (trackCollection[trackKey].BitRate > tracks[i].BitRate)
                            {
                                IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)tracks[i];
                                numberDuplicateFound++;
                                tracksToRemove.Add(tracks[i]);
                            }
                            else
                            {
                                IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)tracks[i];
                                trackCollection[trackKey] = fileTrack;
                                numberDuplicateFound++;
                                tracksToRemove.Add(tracks[i]);
                            }                            
                        }
                    }
                }                                
            }

            SetupProgress(tracksToRemove.Count);

            for (int i = 0; i < tracksToRemove.Count; i++)
            {
                IITFileOrCDTrack track = (IITFileOrCDTrack)tracksToRemove[i];
                UpdateLabel("Removing " + track.Name);
                IncrementProgress();
                AddTrackToList((IITFileOrCDTrack)tracksToRemove[i]);

                if (checkBoxRemove.Checked)
                    track.Delete();
            }

            UpdateLabel("Checked " + numberChecked.ToString() + " tracks and " + numberDuplicateFound.ToString() + " duplicate tracks found.");
            SetupProgress(1);
        }

        private void FindDeadTracks()
        {
            //create a reference to iTunes
            var iTunes = new iTunesAppClass();

            //get a reference to the collection of all tracks
            IITTrackCollection tracks = iTunes.LibraryPlaylist.Tracks;

            int trackCount = tracks.Count;
            int numberChecked = 0;
            int numberDeadFound = 0;

            //setup the progress control
            SetupProgress(trackCount);

            for (int i = trackCount; i > 0; i--)
            {
                if (!_shouldStop)
                {
                    IITTrack track = tracks[i];
                    numberChecked++;
                    IncrementProgress();
                    UpdateLabel("Checking track # " + numberChecked.ToString() + " - " + track.Name);
                    
                    if (track.Kind == ITTrackKind.ITTrackKindFile)
                    {
                        IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)track;                        

                        //if the file doesn't exist, we'll delete it from iTunes
                        if (fileTrack.Location == String.Empty)
                        {
                            numberDeadFound++;
                            AddTrackToList(fileTrack);

                            if (checkBoxRemove.Checked)
                            {
                                fileTrack.Delete();
                            }
                        }
                        else if (!System.IO.File.Exists(fileTrack.Location))
                        {
                            numberDeadFound++;
                            AddTrackToList(fileTrack);

                            if (checkBoxRemove.Checked)
                            {
                                fileTrack.Delete();
                            }
                        }
                    }
                }
            }

            UpdateLabel("Checked " + numberChecked.ToString() + " tracks and " + numberDeadFound.ToString() + " dead tracks found.");
            SetupProgress(1);
        }

        #region Button and Form Handlers
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            _shouldStop = true;
            buttonCancel.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "";
            buttonCancel.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _shouldStop = false;
            buttonCancel.Enabled = true;
            listView1.Items.Clear();

            worker = new Thread(FindDeadTracks);
            worker.Start();
        }
        #endregion

        #region Delegate Callbacks
        //delagates for thread-safe access to UI components
        delegate void SetupProgressCallback(int max);
        delegate void IncrementProgressCallback();
        delegate void UpdateLabelCallback(string text);
        delegate void CompleteOperationCallback(string message);
        delegate void AddTrackToListCallback(IITFileOrCDTrack fileTrack);

        private void IncrementProgress()
        {
            if (progressBar1.InvokeRequired)
            {
                var cb = new IncrementProgressCallback(IncrementProgress);
                Invoke(cb, new object[] { });
            }
            else
                progressBar1.PerformStep();
        }

        private void UpdateLabel(string text)
        {
            if (label1.InvokeRequired)
            {
                var cb = new UpdateLabelCallback(UpdateLabel);
                Invoke(cb, new object[] { text });
            }
            else
                label1.Text = text;
        }

        private void CompleteOperation(string message)
        {
            if (label1.InvokeRequired)
            {
                var cb = new CompleteOperationCallback(CompleteOperation);
                Invoke(cb, new object[] { message });
            }
            else
                label1.Text = message;
        }

        private void AddTrackToList(IITFileOrCDTrack fileTrack)
        {
            if (listView1.InvokeRequired)
            {
                var cb = new AddTrackToListCallback(AddTrackToList);
                Invoke(cb, new object[] { fileTrack });
            }
            else
                listView1.Items.Add(new ListViewItem(new string[] { fileTrack.Name, fileTrack.Artist, fileTrack.Location, fileTrack.BitRate.ToString() }));
        }

        private void SetupProgress(int max)
        {
            if (progressBar1.InvokeRequired)
            {
                SetupProgressCallback cb = new SetupProgressCallback(SetupProgress);
                Invoke(cb, new object[] { max });
            }
            else
            {
                progressBar1.Maximum = max;
                progressBar1.Minimum = 1;
                progressBar1.Step = 1;
                progressBar1.Value = 1;
            }
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            _shouldStop = false;
            buttonCancel.Enabled = true;
            listView1.Items.Clear();

            worker = new Thread(RemoveDuplicates);
            worker.Start();
        }
    }
}
