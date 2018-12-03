import win32com.client
import os
import math
import tkinter as tk
from tkinter import *

class Application(tk.Frame):
    def __init__(self, master=None):
        self.lastTime = 0
        self.playlists = []
        for source in itunes.Sources:
            if source.Name == "Library":
                for playlist in source.Playlists:
                    self.playlists.append(playlist)
                    if playlist.Name == "Library":
                        self.library = playlist

        tk.Frame.__init__(self, master)
        self.grid()
        self.createWidgets()
        self.getTime()

    def createWidgets(self):
        # Song info stuff
        if itunes.PlayerState == 0:
            self.stateLabel = tk.Label(self, text="Paused")
        else:
            self.stateLabel = tk.Label(self, text="Playing")
        try:
            self.songLabel = tk.Label(self, text=itunes.CurrentTrack.Name)
        except:
            self.songLabel = tk.Label(self, text="No song selected")

        self.songSlider = Scale(self, orient=HORIZONTAL, showvalue=0, command=self.moveSlider)
        self.songTime = Label(self, text="0:00")
        self.startTimeLabel = Label(self, text="0:00")
        self.stopTimeLabel = Label(self, text="0:00")

        rowStart = 0
        columnStart = 0
        self.stateLabel.grid(row=rowStart,column=columnStart,columnspan=2)
        self.songLabel.grid(row=rowStart+1,column=columnStart,columnspan=2)
        self.songTime.grid(row=rowStart+2,column=columnStart,columnspan=2,)
        self.startTimeLabel.grid(row=rowStart+3,column=columnStart,columnspan=2,sticky='w')
        self.stopTimeLabel.grid(row=rowStart+3,column=columnStart,columnspan=2,sticky='e')
        self.songSlider.grid(row=rowStart+3,column=columnStart,columnspan=2)

        # Control stuff
        self.controlLabel = tk.Label(self, text='Controls')
        self.playButton = tk.Button(self, text='Play', command=self.play, width=10)
        self.pauseButton = tk.Button(self, text='Pause', command=self.pause, width=10)
        self.nextButton = tk.Button(self, text='Next', command=self.nextSong, width=4)
        self.prevButton = tk.Button(self, text='Prev', command=self.prevSong, width=4)
        self.muteButton = tk.Button(self, text='Mute', command=self.mute, width=10)
        self.quitButton = tk.Button(self, text='Quit', command=self.quit, width=10)

        rowStart = 4
        columnStart = 0
        self.controlLabel.grid(row=rowStart,column=columnStart)
        self.playButton.grid(row=rowStart+1,column=columnStart)
        self.pauseButton.grid(row=rowStart+2,column=columnStart)
        self.nextButton.grid(row=rowStart+3,column=columnStart,sticky='e')
        self.prevButton.grid(row=rowStart+3,column=columnStart,sticky='w')
        self.muteButton.grid(row=rowStart+4,column=columnStart)
        self.quitButton.grid(row=rowStart+5,column=columnStart)

        # Playsong stuff
        self.selectedPlaylist = tk.StringVar(self)
        self.selectedPlaylist.set(self.playlists[0].Name)

        self.playSongLabel = tk.Label(self, text='Play Song')
        self.playSongTextbox = tk.Entry(self)
        self.playSongPlaylists = tk.OptionMenu(self, self.selectedPlaylist, *[playlist.Name for playlist in self.playlists])
        self.playSongListbox = tk.Listbox(self, height=3)
        self.playSongSearchButton = tk.Button(self, text='Search', command=self.listSongs, width=7)
        self.playSongPlayButton = tk.Button(self, text='Play', command=self.playSong, width=7)

        rowStart = 4
        columnStart = 1
        self.playSongLabel.grid(row=rowStart,column=1)
        self.playSongTextbox.grid(row=rowStart+1,column=1)
        self.playSongPlaylists.grid(row=rowStart+2,column=1)
        self.playSongListbox.grid(row=rowStart+3,column=1,rowspan=2)
        self.playSongSearchButton.grid(row=rowStart+5,column=1,sticky='w')
        self.playSongPlayButton.grid(row=rowStart+5,column=1,sticky='e')

    def getTime(self):
        try:
            time = itunes.PlayerPosition
            self.lastTime = time
            self.songTime.configure(text=self.formatTime(time))
            self.songSlider.set(time)
        except:
            self.songTime.configure(text="0:00")

        self.after(100, self.getTime)

    def moveSlider(self, val):
        if int(val) != int(itunes.PlayerPosition):
            itunes.PlayerPosition = val

    def nextSong(self):
        itunes.NextTrack()

    def prevSong(self):
        itunes.PreviousTrack()

    def listSongs(self):
        names = [playlist.Name for playlist in self.playlists]
        songs = self.playlists[names.index(self.selectedPlaylist.get())].Search(self.playSongTextbox.get(), 5)
        self.playSongListbox.delete(0,tk.END)
        try:
            for song in songs:
                self.playSongListbox.insert(tk.END, song.Name)
        except:
            messagebox.showerror("Error", "No songs with the name " + self.playSongTextbox.get() + " in " + self.playlists[names.index(self.selectedPlaylist.get())].Name)

    def playSong(self):
        names = [playlist.Name for playlist in self.playlists]
        self.playlists[names.index(self.selectedPlaylist.get())].PlayFirstTrack()
        song = self.playlists[names.index(self.selectedPlaylist.get())].Search(self.playSongListbox.get(self.playSongListbox.curselection()), 5).Item(1)
        song.Play()

    def play(self):
        itunes.Play()

    def pause(self):
        itunes.Pause()

    def mute(self):
        if itunes.Mute:
            itunes.Mute = False
            self.muteButton.configure(text='Mute')
        else:
            itunes.Mute = True
            self.muteButton.configure(text='Unmute')

    def quit(self):
        itunes.Quit()
        exit()

    def formatTime(self, sec):
        minutes = math.floor(sec/60)
        sec = sec%60
        if sec < 10:
            sec = '0' + str(sec)
        return str(minutes) + ':' + str(sec)

class Events:
        def OnPlayerPlayEvent(self, song):
            app.songLabel.configure(text=itunes.CurrentTrack.Name + " from " + itunes.CurrentPlaylist.Name)
            app.stateLabel.configure(text="Playing")

            app.songSlider.configure(to=itunes.CurrentTrack.Duration)
            app.songSlider.set(itunes.PlayerPosition)

            app.stopTimeLabel.configure(text=app.formatTime(itunes.CurrentTrack.Duration))

        def OnPlayerStopEvent(self, song):
            app.stateLabel.configure(text="Paused")

itunes = win32com.client.DispatchWithEvents("iTunes.Application", Events)
app = Application()
app.master.title('iTunes Simple UI')
app.mainloop()