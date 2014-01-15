# outlook-notify-my-android

This repository contains some very basic VBA (the first I've written in probably 15 years :P) to forward notifications from Outlook to Notify My Android for those who don't want to sign up to the potentially onerous permissions of the official Outlook client for Android.

## Code

### Module1

*NotifyMyAndroidMessageRule* - a sub that accepts a mail item and creates a Notify My Android message containing a summary
*NotifyMyAndroid* - a public sub for sending Notify My Android messages
*URLEncode* - a public function ripped off the internet (apologies for lack of accreditation, can't remember which site I found it on) for performing URL encoding as VBA doesn't seem to support it natively

### ThisOutlookSession

*Application_Reminder* - trigger to send a Notify My Android notification in response to Outlook reminders.  Code based on this: http://www.outlookcode.com/d/code/sendreminder.htm

## Installation

The relevant code needs to be put into the relevant modules in Outlook Visual Basic (Developer -> Visual Basic in Outlook 2010).  The project needs to reference WinHttp.  The Notify My Android apikey and app name need to be set in the code.  Look for the following:

+ INSERT YOU NOTIFY MY ANDROID APIKEY HERE
+ INSERT YOUR APP NAME HERE

This is a bit ugly but my VBA knowledge does not extend to parameterising these.
