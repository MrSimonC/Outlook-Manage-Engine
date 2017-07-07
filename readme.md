# Outlook Mail Traverser for Manage Engine
Made for North Bristol Trust Back Office Team, whilst working as a Senior Clinical Systems Analyst.

## Description
Traverse an Outlook shared mailbox, searching for a regular expression pattern in emails identifying a local reference number for the [Manage Engine Helpdesk System](https://www.manageengine.com), confirm reference number exists via Service Desk Plus API, then forward onto a monitored email address, putting the local reference number in ##s.

## Background
If you have a supplier with a rigid bespoke call logging system which can't invoke the ##s in email subject replies (which is what your local helpdesk system uses to identify email call responses), then have the supplier send into one shared mailbox, traverse each incoming email for the local reference and forward on to your Manage Engine system whilst checking the local ref number validity.

Example:  
Imagine your company owns a Manage Engine helpdesk system which identifies call responses with the notation ##local_reference## in the subject.  
e.g. "Re: Your previous call ##12345##"  

However, if they send in  
e.g. "Re: Your previous call 12345"  
This will create a new call in Manage Engine.

Instead:
- have them email "externalsupplier@company.com"
- run this program on a PC with access to that mailbox
- each email will be put checked for a pattern, forwarded onto the main manage engine email address, with personal signature removed, and with a line "to respond to the supplier, use this email: supplier@suppliercompany.com"

## Prerequisite
On the server PC, install "Visual C++ Redistributable for Visual Studio 2015 x86.exe" (on 32-bit, or x64 on 64-bit) which allows Python 3.5 dlls to work, found here:
https://www.microsoft.com/en-gb/download/details.aspx?id=48145

## Installation and Running
Make a folder in "C:\Program Files\Outlook SDPlus" (or another location of you choice) and put all the following files in there:
- outlook_sdplus.exe _(found in the "binaries" folder)_

In Windows Task Scheduler, Import Task, choose the .xml file, then change the "Run only when user is logged on" username to your own.
OR create a new task with the following attributes:
- General: Run only when user is logged on
- Trigger: at 8am everyday, repeat every 15 minutes indefinitely
- Actions: Start a program: "C:\Program Files\Outlook SDPlus\outlook_sdplus.exe"
- Settings:
    - Allow ask to be run on demain
    - Stop the task if it runs longer than: 3 days
    - If the running task does not end when requested, force it stop
    - If the task is already running, then the following rule applies: Stop the existing instance

## SDPlus API
This program communicates with the sdplus API via an sdplus api technician key which can be obtained via the sdplus section: Admin, Assignees, Edit Assignee (other than yourself), Generate API Key.
This program will look for an API key in a windows variable under name "SDPLUS_ADMIN". You can set this on windows with:
`setx SDPLUS_ADMIN <insert your own SDPLUS key here>`
in a command line.

## Outlook
Ensure your Outlook inbox doesn't fill up - else this will create error messages on your desktop when the program runs.
This program looks for the IT Third Party Response inbox via 'RM1048' - the 'Alias' of IT Third Party Response.

## Slack
This program communicates with slack API via an API Token which can be obtained by: https://it-nbt.slack.com/services/B1FCFC4RL (or Browse Apps  > Custom Integrations  > Bots  > Edit configuration)
This program will look for an API key in a windows variable under name "SLACK_LORENZOBOT". You can set this on windows with:
`setx SLACK_LORENZOBOT <insert your own SDPLUS key here>`
in a command line.

Written by:
Simon Crouch, late 2016 in Python 3.5