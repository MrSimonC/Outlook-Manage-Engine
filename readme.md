# Outlook Mail Traverser for Manage Engine
## Background
Traverse an Outlook shared mailbox,  searching for a regular expression pattern in emails identifying a local reference number for the [Manage Engine Helpdesk System](https://www.manageengine.com), confirm local reference exists via API, then forward onto a monitored email address, putting the local reference number in ##s.

If you have a supplier with a rigid bespoke call logging system which can't invoke the ##s in email subject replies (which is what your local helpdesk system uses to identify email call responses), then have the supplier send into one shared mailbox, traverse each incoming email for the local reference and forward on to your Manage Engine system whilst checking the local ref number validity.

## Usage
Example:</br>
Imagine your company owns a Manage Engine helpdesk system which identifies call responses with the notation ##local_reference## in the subject.</br>
e.g.  Call reference: 12345</br>
<b>Email Subject: Re: Your previous call ##12345##</br></b>
However, they send in</br>
<b>Re: Your previous call 12345</br></b>
This will create a new call in Manage Engine. Instead, have them email "externalsupplier@company.com". Run this program on a PC with access to that mailbox, and each email will be put checked for a pattern, forwarded onto the main manage engine email address, with personal signature removed, and with a line "to respond to the supplier, use this email: supplier@suppliercompany.com"

Written with Python 3.5

### Running
Set the needed Service Desk Plus API key environment variable with `setx SDPLUS_API <Enter your API key here>`

In "binaries" folder, v.91 communicates with slack API via an API Token which can be obtained by: https://it-nbt.slack.com/services/ (or Browse Apps  > Custom Integrations  > Bots  > Edit configuration).  
This program will look for an API key in a windows variable under name "SLACK_LORENZOBOT". You can set this on windows with `setx SLACK_LORENZOBOT <insert your own SDPLUS key here>` in a command line.

Use `pyinstaller outlook_manage_engine.spec` to create a binary executable, then setup a simple windows task scheduler to call the binary every 5 minutes.
