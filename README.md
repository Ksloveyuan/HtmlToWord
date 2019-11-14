# HtmlToWord

The project provides a http service to cover html to word by Microsoft.Office.Interop.Word. And the http server is written by WCF webHttpBinding.

# Prequisites
- .net framework 4.7
- Office word 2013
- Visual Studio 2017

# How to build

First, use nuget to install all neeed package.

Second, build HtmlToWord.ConsoleHost.

If you need to host the service by Windows Service, please build HtmlToWord.WindowsService.

# The request workflow

1. user sends html+css text to server
2. save the html+css text to a html file
3. use word to open html file, then convert to word file
4. return a download link
5. add a url mapping in nginx, then the download part will be handled by nginx
