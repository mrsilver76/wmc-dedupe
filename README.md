# wmc-dedupe
A script which will identify all duplicate Windows Media Center recorded television shows (in either WTV or DVR-MS formats) and then either move them into a folder for duplicates or delete them. Recordings that are sitting in the duplicates folder can be automatically deleted after a certain number of days.

## Why use wmc-dedupe?

Although the series link functionality within Windows Media Center is pretty good,
unfortunately due to either poor programming by Microsoft or <a href="http://experts.windows.com/frms/windows_entertainment_and_connected_home/f/114/t/100718.aspx">poor EPG data from the
third parties that Microsoft use</a> (or possibly both) it means that setting up a series
recording of a show can often result many duplicate recordings. This is because
Windows Media Center often just ends up recording every single time the show is aired.

Ideally Microsoft (or their EPG provider) would fix the problem but, since Windows Media Center is now deprecated, this is unlikely to happen. As a result, wmc-dedupe temporarily "fixes" the problem by deleting duplicate recordings and saves you from drowning in repeats of a show.

## Features

*  Highly configurable command line based program for running as a one off or as a scheduled task.
*  Looks at WTV or DVR-MS files (with extensions wtv, dvr-ms or dvrms).
*  Options to exclude examing WTV or DVR-MS files.
*  Can delete duplicate shows or move them into a user defined duplicates folder.
*  Only the larger duplicate recording is kept (meaning that HDTV or longer shows are prioritised)
*  Can delete shows from the duplicates folder after they are older than a certain number of days.
*  Test mode which doesn't delete or move files.
*  Seven days worth of logs kept.
*  Three levels of logging, from minimal to very verbose.
*  Ability to pause after running to view the output.
*  Can use the "Public Recorded TV" path or any other location.
*  Automatically detects if a recording is occuring. Can be forced to ignore this if required.
*  Pre-loading of meta-data for fast duplicate matching over a slow connection (such as a network).
*  Duplicates identifed even if they have differing punctuation, capitalisation or accents.
*  Descriptions with brackets (either square or normal) are handled correctly when comparing files.
*  Common starting descriptions (e.g. "repeat", "premier" or "new") are handled when comparing files.
*  Descriptions where the sub-title has been incorrectly placed into the description are handled when comparing files.
*  Sanity checking to ensure that incorrect command line options don't destroy your recorded TV.

## Requirements

In order to use this program you need the following:

*   A computer running Microsoft Windows and Windows Media Center.
*   Some recorded television shows in WTV or DVR-MS format.
*   Knowledge of running command line applications.

This program is not recommended for people who are not comfortable with the workings of Microsoft Windows, command line applications and the scheduling of tasks.

## Installation and usage

1.  Download the latest version from https://github.com/mrsilver76/wmc-dedupe/releases
2.  Extract the files and copy `wmc-dedupe-1.5.vbs` into any folder on your computer. One possible option is the Documents folder.
3.  Double clicking on the program will pop up a message displaying the command line options.
4.  To run the program from the command line, you should enter the following command from within a DOS windows:  
    `cscript wmc-dedupe-1.5.vbs [options]` where  `[options]` are the possible options detailed in the next section.
5.  All logs are stored within the Application Data folder. The easiest way to access is to enter the following either in the start menu search, from the Run command or in the address of the Explorer window: `%appdata%\wmc-dedupe` This will open the browser window at the location of the logs.
6.  You can call this program automatically using Windows Task Scheduler. For more details, see the section later.
7.  You can call this program by double-clicking on an icon. To do this you need to right-click on the program and select "Create Shortcut". When the shortcut appears, right-click on that, select "Properties" and add the command line options to the end of the section entitled "Target".

## Command line options overview

    cscript wmc-dedupe-1.5.vbs [/M | /D] [/R | <tv path>] [/E:x] [dup path] [/T] [/F] [/S] [/NW|/ND] [/P] [/V | /VV]
    
        No args     Display help. This is the same as typing /?.
        /?          Display help. This is the same as not typing any options.
        /M          Move duplicate recordings into another folder.
        /D          Delete duplicate recordings.
        /R          Look at shows in the Public 'Recorded TV' location.
        <tv path>   The path to Recorded TV files. Required unless /R is used.
        /E:x        Erase files in the duplicates folder older than x days.
        [dup path]  Path to duplicates folder. Required with /M or /E.
        /T          Test mode. Don't move or delete any shows.
        /F          Force execution even if something is being recorded.
        /S          Silent. Don't display anything on the screen.
        /NW         Ignore WTV files.
        /ND         Ignore DVR-MS files.
        /P          Pause after running.
        /V          Verbose mode. Log additional information during execution.
        /VV         Very Verbose mode. Log even more information. Implies /V.

**Please note!** The `[` and `]` denote optional arguments and the `<` and `>` denote mandatory arguments, both of which should not be used as part of the command line. If you are unsure, please refer to the examples further down this README.

## Command line options details

1.  If you do not prefix the command with `cscript` then the program will run it its own popup window.
2.  Options are prefixed with a `/`. If you prefer, you can use `-` (for example, `-R` and `/R` are the same).
3.  If your choice of command line options means that you have to supply two paths, then the path to duplicates _**always comes after**_ the Recorded TV path.
4.  When supplying paths, if there are any spaces in the path then it _**must**_ be enclosed in speech marks. To be safe, it is recommended to always use speech marks.
5.  If you try and supply the same folder name for both the location of recordings and the duplicates folder then the program will warn you and exit.
6.  If you use the standard 'Recorded TV' location for storing (or erasing) duplicates then the program will warn you and exit. This is to prevent you from accidentally deleting all your recordings.

### Move duplicate recordings into another folder (/M)

Tells the program that duplicate recordings from one folder should be moved into another folder. To use this command you must supply the path to the 'Recorded TV' location (or use `/R`) and the path where to store duplicates.

If both paths point to the same place, then the program will display an error

### Delete duplicate recordings (/D)

Tells the program that duplicate recordings should be deleted. To use this command you must supply the path to the 'Recorded TV' location (or use `/R`).

You cannot use `/D` and `/M` at the same time.

### Look at shows in the Public 'Recorded TV' location (/R)

Tells the program to assume that recorded television shows are located in the public 'Recorded TV' folder. This varies depending on how Windows was installed and the program is clever enough to cope if you have this in a non-standard location or drive (for most people running Windows 7 it is `C:\Users\Public\Recorded TV`). If you use this command then you do not need to supply a path to the 'Recorded TV' folder.

### Erase files in the duplicates folder older than x days (/E:x)

Tells the program that any recordings in the duplicates folder that are older than x days should be deleted. For example, using `/E:5` means erase files that are older than 5 days. You cannot set this value any lower than 1.

In order to use this command you must supply the path to where the duplicates are stored. If your path points to the public 'Recorded TV' location then an error will occur to prevent you from accidently deleting old recordings.

### Test mode. Don't move or delete any shows (/T)

Runs the program as normal but does not delete or move any files. Useful if you want to see what will happen before you run the program properly.

### Force execution even if something is being recorded (/F)

By default, the program will not run if something is being recorded. This is to prevent duplicate checking on a program which is only partially recorded and any errors that occur if a file is attempted to be moved which is being used for a recording.

If you use this command then the program will be forced to continue to run, even if something is recording.

### Silent. Don't display anything on the screen (/S)

Ensures that nothing is displayed to the screen whilst the program is being run. This option is largely useless on the basis that Windows Media Center will always sit on top of anything else that is running.

### Ignore WTV files (/NW)

If you use this command then files ending in WTV will be ignored when looking for duplicates. If you use this option and also include `/ND` then no files will be scanned at all. If you don't want to scan any files, then the better method is to not include `/M` or `/D` in your command line.

### Ignore DVR-MS files (/ND)

If you use this command then files ending in DVR-MS or DVRMS will be ignored when looking for duplicates. If you use this option and also include `/NW` then no files will be scanned at all. If you don't want to scan any files, then the better method is to not include `/M` or `/D` in your command line.

### Pause after running (/P)

If you use this command then after the program has finished running, you will be presented with a prompt to press Enter before it finishes. This is not recommended if you plan to run the command from a batch file or from a scheduled task.

### Verbose mode. Log additional information during execution (/V)

Logs (and displayed to the screen, unless `/S` has been used) even more information about what the program is doing during running. Useful for debugging and the curious. Be warned that the extra debugging information will slow down the running of the program significantly and generate logs which are easily a megabyte or more in size each time.

If you plan to send these log files by email, then using zip to compress them is highly recommended.

### Very Verbose mode. Log even more information (/VV)

Like verbose mode but also includes details of the title, sub-title and description of every file being compared. This is only useful for debugging, is even slower than using `/V` and will generate logs which are easily in the tens of megabytes.

If you plan to send these log files by email, then using zip to compress them is essential.

## Example usage

Below are a couple of recommended command line uses:

    cscript wmc-dedupe-1.5.vbs /M /R /E:14 "C:\Users\Public\Documents\Duplicates"

Move (`/M`) duplicate recordings in the public 'Recorded TV' folder (`/R`) into `C:\Users\Public\Documents\Duplicates`. After this has happened, look in `C:\Users\Public\Documents\Duplicates` and delete any recordings which are over a fortnight old (`/E:14`).

    cscript wmc-dedupe-1.5.vbs /D /R

Delete (`/D`) any duplicate recordings in the public 'Recorded TV' folder (`/R`).

    cscript wmc-dedupe-1.5.vbs /E:5 "D:\Duplicate Recordings" /F

Look in the `D:\Duplicate Recordings` folder and delete any recordings that are over 5 days old (`/E:5`). Ignore the fact that something might currently be recordings (`/F`).

## Scheduled task

The recommended method to run this program is via a scheduled task. [This Microsoft website](http://windows.microsoft.com/en-US/windows7/schedule-a-task) explains how to create one in Windows 7. Some key points to know:

1.  Using "Create Task..." instead of "Create Basic Task..." will give you more flexibility.
2.  The trigger could be one or more times a day (based around your viewing habits) or when the computer turns on.
3.  In Actions: The "Program/Script" should point to wmc-dedupe. The "Add arguments" should be all the command line options. You do not need to fill out "Start in"
4.  Bear in mind that if you don't use /F then the program will not run when something is recording on the television.

## Recommended setup

There is no recommended setup. If you are interested, my HTPC runs a batch file twice a day (at midday and 5pm) which moves recorded TV shows that are movies into a seperate folder (so they don't clutter up the 'Recorded TV' section) and then the following:

    cscript wmc-dedupe-1.5.vbs /M /R /E:30 "C:\Users\Public\Documents\Duplicates"

This moves all duplicates from the 'Recorded TV' section into a new folder and then deletes anything in that folder older than 1 month. This means that any shows which are accidently flagged as a duplicate can be rescued within that time.
