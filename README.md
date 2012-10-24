TVTorrentMover
==============

A WSH/vbscript implementation that moves tv shows from the torrent download folder to shares on the network. It takes tv shows with their torrent filename and copies them.

the way I run this is on my PVC, which runs XBMC on Vista, is to first set cscript as the default scripting host, and create a windows task scheduler job pointing to the script to run every so often (once or twice a day). Then grab the library updater plugin for XBMC and configure that to run an hour or so after the task scheduler, to allow it time to copy files.

You have to edit 2 lines in the script: the "root" folder (line 8) which is where to "monitor" for changes (in my case a qnap nas box), and line 17, an array of shares where the tv shows are stored (e.g. a parent folder containing folders for each tv show) - these can be on different boxes or shares on the same box, etc.

This script will do it's best to find matches for tv shows even if the tv show name and folder name aren't the same (e.g. if you have "My TV Show (2012)" as a folder and the download name is "my.tv.show.(2012).s01e03.release.res.mp4"

the tvshow season/episode format must be S00E00 not 0x00 as some do it. I'm happy to fix the files that don't copy manually so I didn't code this in.

no logging, no interface, no message telling you what didn't copy. want it? fork this and give it a go yourself.