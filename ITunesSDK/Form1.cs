using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTunesLib;
using System.Diagnostics;

namespace ITunesSDK
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Threading.Thread thread = new System.Threading.Thread(findMissingArtTracks);
            thread.Start();
        }

        private void findMissingArtTracks()
        {
            /** This code makes use of the ITunes SDK to find tracks missing the artwork (cd cover)
            and add those tracks to a playlist, so I could later go add the artwork to them
            Remember! to add a reference to iTunes <version #> Type Library 
            Adding the reference on Visual Studio- Go to menu Project>Add Reference.. >> COM and search for 'iTunes <version #> Type Library'
            <version #> = 1.13 for me.
            This COM is added to your computer when you install iTunes. If you can't find it re-installing itunes might help.
            */

            //Create the itunes object
            iTunesLib.iTunesApp oItunes = new iTunesLib.iTunesApp();
            //optional: minimize itunes
            oItunes.BrowserWindow.Minimized = true;

            iTunesLib.IITSourceCollection sources = oItunes.Sources;

            //get all the playlists from itunes' library
            iTunesLib.IITPlaylistCollection playlists = sources.ItemByName["Library"].Playlists;
            string libName = "Missing Artwork";
            //get playlist with specific name
            iTunesLib.IITPlaylist lib = playlists.ItemByName[libName];

            //delete playlist if it already exists
            //note: itunes allows several playlists with same name, this aprach only deletes one of them
            //create a while loop to delete all of them
            if (lib != null) lib.Delete();

            //create a empty playlist
            lib = oItunes.CreatePlaylist(libName);

            //get all the tracks in the itunes' library
            iTunesLib.IITLibraryPlaylist library = oItunes.LibraryPlaylist;
            //using 'var' to show we don't need to remember the object type.
            var tracks = library.Tracks;
            int counter = 0;
            showMessage("Processing..." + libName);
            foreach (IITTrack track in tracks)
            {
                //add to the playlist tracks that don't have artwork
                //we can use the same technique to find other proprieties, song's name, artist, album, etc. 
                if (track.Artwork.Count == 0)
                {
                    counter++;
                    Debug.WriteLine("Adding: {0} ", track.Name);
                    /* IITPlaylist doesn't have the AddTrack so we need to cast to UserPlayList
                    I could have created another variable outside the loop for performance like:
                    IITUserPlaylist userLib = (IITUserPlaylist) objApp.CreatePlaylist(libName);
                    but wanted this way to cast and assign 
                     */
                    (lib as IITUserPlaylist).AddTrack(track);
                }
            }

            string result = String.Format("Total: {0} Track(s) found.", counter);
            Debug.WriteLine(result);
            showMessage(result);
            //oItunes.Quit();
            //oItunes = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Threading.Thread newThread = new System.Threading.Thread(deleteDeadTracks);
            newThread.Start();
        }

       
        private void deleteDeadTracks()
        {
            iTunesLib.iTunesApp objApp = new iTunesLib.iTunesApp();
            iTunesLib.IITLibraryPlaylist objLibrary = objApp.LibraryPlaylist;

            var colTracks = objLibrary.Tracks;
            List<string> songsToDelete = new List<string>();

            //using invoke to show message on a diff thread
            showMessage("Processing... dead tracks");
            foreach (IITTrack track in colTracks)
            {
               // Debug.WriteLine("runnung... "+track.Name);
                iTunesLib.IITFileOrCDTrack objSong = (iTunesLib.IITFileOrCDTrack)track;
                
                if (String.IsNullOrEmpty(objSong.Location))
                {
                    Debug.WriteLine("name: " + objSong.Name);
                    songsToDelete.Add(objSong.Name);

                }

              //System.Threading.Thread.Sleep(50);
            }

            foreach (string song in songsToDelete)
            {
                System.Threading.Thread.Sleep(1000);
                var objSong = colTracks.ItemByName[song];
                objSong.Delete();
            }

            string result = String.Format("Total: {0} Track(s) Deleted.", songsToDelete.Count);
            Debug.WriteLine(result);
            showMessage(result);
            objApp = null;
        }
        delegate void formAccessDelegate(string text);
        public void showMessage(string text)
        {
            if (resultLb.InvokeRequired)
            {
                formAccessDelegate del = new formAccessDelegate(showMessage);
                resultLb.Invoke(del, text);
            }
            else
            {
                resultLb.Text = text;
            }
        }
    }
}
