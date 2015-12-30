using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using iTunesLib;

namespace You
{
    /// <summary>
    /// プログラムのエントリポイントを提供します。
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// プログラムを開始します。
        /// </summary>
        [STAThread]
        public static void Main()
        {
            try
            {
                iTunesApp iTunes = null;
                try
                {
                    iTunes = new iTunesApp();

                    IITSource librarySource = null;
                    try
                    {
                        librarySource = iTunes.LibrarySource;

                        IITPlaylistCollection playlistCollection = null;

                        try
                        {
                            playlistCollection = librarySource.Playlists;

                            var playlistList = new List<IITPlaylist>();

                            try
                            {
                                foreach (IITPlaylist playlist in playlistCollection)
                                {
                                    playlistList.Add(playlist);
                                }

                                playlistList
                                    .AsParallel()
                                    .Select(x => x as IITUserPlaylist)
                                    .Where(x =>
                                        x != null &&
                                        x.Smart == false &&
                                        x.SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone)
                                    .ForAll(x => x.Delete());
                            }
                            finally
                            {
                                foreach (var playlist in playlistList)
                                {
                                    Marshal.ReleaseComObject(playlist);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(playlistCollection);
                        }
                    }
                    finally
                    {
                        if (librarySource != null)
                        {
                            Marshal.ReleaseComObject(librarySource);
                        }
                    }
                }
                finally
                {
                    if (iTunes != null)
                    {
                        Marshal.ReleaseComObject(iTunes);
                    }
                }

                MessageBox.Show("completed.");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }
    }
}