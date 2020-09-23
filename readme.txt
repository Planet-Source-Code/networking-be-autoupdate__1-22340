AutoUpdate - By Cakkie

Usage

au.exe /url <url> [/app <app>] [/tmo <tmo>] [/ret <ret>]

/url <url>
Specifies the url where the version file can be found

/app <app>
If specified, the application to be started when the updating is completed

/tmp <tmo>
The connection timeout in seconds, if not specified, the default is used. default = 10

/ret <ret>
The number of retrys when a file can't be retrieved (eg when the server is busy)
The default is 2

Example
au.exe /url http://www.myserver.com/myfolder/version.dat /app myapp.exe

the version file has following format
"from","to",#date#

from: the url to the file the original file
to: the path to the file where the file must be saved to
date: a date/time containing the date and time that the file was changed

Example
"http://www.myserver.com/myfolder/myfile.dat","##apppath##\myfile.dat",#01/01/2001 7:25:00#

If TO contains ##apppath##, ##apppath## is replaced by the path of au.exe