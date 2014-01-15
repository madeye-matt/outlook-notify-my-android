# outlook-notify-my-android

This repository contains some very basic VBA (the first I've written in probably 15 years :P) to forward notifications from Outlook to Notify My Android for those who don't want to sign up to the potentially onerous permissions of the official Outlook client for Android.

It consists of 2 parts: 

## Module1

*NotifyMyAndroidMessageRule* - a sub that accepts a mail item and creates a Notify My Android message containing a summary
*NotifyMyAndroid* - a public sub for sending Notify My Android messages
*URLEncode* - a public function ripped off the internet (apologies for lack of accreditation, can't remember which site I found it on) for performing URL encoding as VBA doesn't seem to support it natively

## ThisOutlookSession

*Application_Reminder* - trigger to send a Notify My Android notification in response to Outlook reminders.  Code based on this: http://www.outlookcode.com/d/code/sendreminder.htm
