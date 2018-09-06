<html>
<head>
</head>

<img align="left" src="Images/ReadMe/header3.png">

This Add-In is used for marking up Excel files with revisions.
There are two common methods of indicating where a revision has changed a drawing that contains a system diagram. The first is the cloud method, were each change is enclosed by a cloud shape.  The second method involves placing a triangle with the revision number next to each effected portion of the drawing.  The cloud method indicates changes from the most recent revision only, whereas the second method indicates all revisions to the drawing because all of the previous revision circles remain on the drawing.

[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://paypal.me/AnthonyDuguid)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE "MIT License Copyright © Anthony Duguid")
[![Latest Release](https://img.shields.io/github/release/Office-projects/Excel-Markup.svg?label=latest%20release)](https://github.com/Office-projects/Excel-Markup/releases)
[![Github commits (since latest release)](https://img.shields.io/github/commits-since/Office-projects/Excel-Markup/latest.svg)](https://github.com/Office-projects/Excel-Markup)

The most complete version is C#. 
<a href="https://github.com/Office-projects/Excel-Markup/issues" target="_blank">
 <img alt="Issues" src="https://img.shields.io/github/issues/Office-projects/Excel-Markup.svg" />
</a>

<h1 align="left">
  <img src="Images/ReadMe/vsto.excel.ribbon.markup.gif" />
</h1>

<br>

## Table of Contents
- <a href="#install">Install</a>
- <a href="#dependencies">Dependencies</a>
- <a href="#glossary-of-terms">Glossary of Terms</a>
- <a href="#functionality">Functionality</a> 
    - <a href="#help">Help</a>
    - <a href="#about">About</a>

<br>

<a id="user-content-install" class="anchor" href="#install" aria-hidden="true"> </a>
## Install
Instructions for installation of VBA and VSTO versions.

>TODO: deploy to Microsoft Office store

### VBA
How to install the VBA version
1. Download the VBA Add-In file [![download VBA](https://img.shields.io/badge/download-VBA-brightgreen.svg)](https://github.com/Office-projects/Excel-Markup/raw/master/VBA/Markup.xlam?raw=true "Download the VBA Add-In").
2. Copy the file to the XLSTART folder on your computer. ```%AppData%\Microsoft\Excel\XLSTART\```
3. Close all open instances of Excel and then launch Excel. The new ribbon should appear.

### VSTO
How to install the VSTO version
1. Download AnthonyDuguid.pfx And Install At Root Level [![download Key](https://img.shields.io/badge/download-Key-brightgreen.svg)](https://github.com/Office-projects/Excel-Markup/blob/master/CS/AnthonyDuguid.pfx?raw=true "Download AnthonyDuguid.pfx And Install At Root Level For VSTO")
2. Download and run the setup.exe file. [![download VSTO](https://img.shields.io/badge/download-VSTO-brightgreen.svg)](https://github.com/Office-projects/Excel-Markup/blob/master/CS/publish/setup.exe?raw=true "Download Setup.exe Install File")

<br>

<a id="user-content-dependencies" class="anchor" href="#dependencies" aria-hidden="true"> </a>
## Dependencies
|Software                                   |Dependency                 |Project                    |
|:------------------------------------------|:--------------------------|:--------------------------|
|[Microsoft Visual Studio Community 2017](https://www.visualstudio.com/vs/whatsnew/)|Solution|VSTO|
|[Microsoft Office Developer Tools](https://visualstudio.microsoft.com/vs/features/office-tools/)|Solution|VSTO|
|[Microsoft Excel 2010 (or later)](https://www.microsoft.com/en-au/software-download/office)|Project|VBA, VSTO|
|[Visual Basic for Applications](https://msdn.microsoft.com/en-us/vba/vba-language-reference)|Code|VBA|
|[Extensible Markup Language (XML)](https://www.rondebruin.nl/win/s2/win001.htm)|Ribbon|VBA, VSTO|
|[Log4Net](https://www.nuget.org/packages/log4net/) |Error Logging |VSTO|
|[ScreenToGif](http://www.screentogif.com/)|Read Me|VBA, VSTO|
|[Snagit](http://discover.techsmith.com/snagit-non-brand-desktop/?gclid=CNzQiOTO09UCFVoFKgod9EIB3g)|Read Me|VBA, VSTO|
|Badges ([Library](https://shields.io/), [Custom](https://rozaxe.github.io/factory/), [Star/Fork](http://githubbadges.com))|Read Me|VBA, VSTO|

<!---|[Sandcastle](https://github.com/EWSoftware/SHFB)|API Documentation|VSTO|--->
<br>

<a id="user-content-glossary-of-terms" class="anchor" href="#glossary-of-terms" aria-hidden="true"> </a>
## Glossary of Terms

| Term                      | Meaning                                                                                  |
|:--------------------------|:-----------------------------------------------------------------------------------------|
| COM |Component Object Model (COM) is a binary-interface standard for software components introduced by Microsoft in 1993. It is used to enable inter-process communication and dynamic object creation in a large range of programming languages. COM is the basis for several other Microsoft technologies and frameworks, including OLE, OLE Automation, ActiveX, COM+, DCOM, the Windows shell, DirectX, UMDF and Windows Runtime.  |
| Cloud |Refers to the joined arcs that make up a shape|
| Hatch |The use of fine, parallel lines drawn closely together, often rapidly drawn, to create the illusion of shade or texture|
| VBA |Visual Basic for Applications (VBA) is an implementation of Microsoft's event-driven programming language Visual Basic 6 and uses the Visual Basic Runtime Library. However, VBA code normally can only run within a host application, rather than as a standalone program. VBA can, however, control one application from another using OLE Automation. VBA can use, but not create, ActiveX/COM DLLs, and later versions add support for class modules.|
| VSTO |Visual Studio Tools for Office (VSTO) is a set of development tools available in the form of a Visual Studio add-in (project templates) and a runtime that allows Microsoft Office 2003 and later versions of Office applications to host the .NET Framework Common Language Runtime (CLR) to expose their functionality via .NET.|
| XML|Extensible Markup Language (XML) is a markup language that defines a set of rules for encoding documents in a format that is both human-readable and machine-readable.The design goals of XML emphasize simplicity, generality, and usability across the Internet. It is a textual data format with strong support via Unicode for different human languages. Although the design of XML focuses on documents, the language is widely used for the representation of arbitrary data structures such as those used in web services.|
<body>

<br>

<a id="user-content-functionality" class="anchor" href="#functionality" aria-hidden="true"> </a>
## Functionality
This Excel ribbon is inserted after the “Home” tab when Excel opens. Listed below is the detailed functionality of this application and its components.  


<a id="user-content-clipboard" class="anchor" href="#clipboard" aria-hidden="true"> </a>
###	Clipboard (Group)

<a id="user-content-paste" class="anchor" href="#paste" aria-hidden="true"> </a>
####	Paste (Menu)
* Paste the contents of the clipboard



<a id="user-content-help" class="anchor" href="#help" aria-hidden="true"> </a>
###	Help (Group)

<a id="user-content-how-to" class="anchor" href="#how-to" aria-hidden="true"> </a>
#### How To… (Button)
* Opens a webpage of the read me documentation

<a id="user-content-api-doc" class="anchor" href="#api-doc" aria-hidden="true"> </a>
#### Report Issue (Button)
* Opens a page to create a new issue for the product

<a id="user-content-settings" class="anchor" href="#settings" aria-hidden="true"> </a>
#### Add-In Settings (Button)

<kbd>
VSTO
<br>
  <img align="left" src="Images/ReadMe/vsto.ribbon.settings.png" />
</kbd>

- Types of VSTO Settings
  - Application Settings
    - These settings can only be changed in the project and need to be redeployed
    - They will appear disabled in the form
  - User Settings
    - These settings can be changed by the end-user
    - They will appear enabled in the form
    
<kbd>
VBA
<br>
  <img align="left" src="Images/ReadMe/vba.ribbon.settings.png" />
</kbd>

- VBA Settings
  - To add a new setting
    ```vbnet
    ThisWorkbook.CustomDocumentProperties.Add _
    Name:="App_ReleaseDate" _
    , LinkToContent:=False _
    , Type:=msoPropertyTypeDate _
    , Value:="31-Jul-2017 1:05pm"
    ```
  - To update a setting
    ```vbnet
    ThisWorkbook.CustomDocumentProperties.Item("App_ReleaseDate").Value = "31-Jul-2017 1:05pm"
    ```
  - To delete a setting
    ```vbnet
    ThisWorkbook.CustomDocumentProperties.Item("App_ReleaseDate").Delete
    ```

<a id="user-content-about" class="anchor" href="#about" aria-hidden="true"> </a>
###	About (Group)
<h1 align="left">
  <img src="Images/ReadMe/ribbon.group.about.png" />
</h1>

<a id="user-content-description" class="anchor" href="#description" aria-hidden="true"> </a>
#### Add-in Name (Label)
* The application name with the version

<a id="user-content-install-date" class="anchor" href="#install-date" aria-hidden="true"> </a>
#### Release Date (Label)
* The release date of the application

<a id="user-content-copyright" class="anchor" href="#copyright" aria-hidden="true"> </a>
#### Copyright (Label)
* The author’s name

</body>
</html>
