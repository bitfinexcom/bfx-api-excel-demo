# BFX() Function Quick Start Guide

This quick start guide will allow you to install and enable the BFX() custom function in your worksheet. A full guide for the function can be found here: 

# Installation

1. Download the BFX Function 1.0.xlam file
2. Copy and paste the BFX Function 1.0.xlam file to the Microsoft Add-ins folder. This is usually found at C:\Users<Username>\AppData\Roaming\Microsoft\Addins

Note: You may need to enable the ability to view hidden files and folders to see the AppData folder under your folder. Instructions to enable this can be found here: https://support.microsoft.com/en-us/windows/view-hidden-files-and-folders-in-windows-10-97fbc472-c603-9d90-91d0-1166d1d9f4b5

# Enabling the add-in in your worksheet

1. In Excel, go to "File" -> "Options" -> "Add-ins".
2. At the bottom of the window that opened, select "Manage: Excel Add-ins" and click on "Go".
3. In the next popup, check the "BFX Function" Add-in and click on "OK".

# Enabling required references

1. Open the VBA editor from the developer tab by clicking on "Visual Basic".
2. In the VBA editor, go to "Tools" -> "References".
3. In the menu that opens, look through the list and place a check at both "Microsoft WinHTTP Services" and "Microsoft Scripting Runtime".
4. Click on "OK" to enable these references in your workbook.

# Ready for use

The BFX() function is now ready for use. In the CheatSheet.xlsx file included in this repository you will find a list of all possible fields that can be retrieved by the add-in.
