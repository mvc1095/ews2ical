ews2ical.php
============

Pulls events from an Exchange Web Service and generate an iCal file for import into another system, eg. Google Calendar, because *anything* is better than using Outlook.

This script is pretty much just sample code and libraries from these two sources:

 - https://github.com/jamesiarmes/php-ews
 - https://github.com/markuspoerschke/iCal

Installation
------------

 1. Download this code
 1. Run `composer install` to fetch the required libraries (see https://getcomposer.org/download/ if you don't already have composer installed)
 1. Place your Office 365 username (eg. *first.last@domain.tld*) in the file `ews2ical.username` in this directory
 1. Place your password in the file `ews2ical.secret` in this directory
 1. Run this script to produce an iCal file as output, and import it into your favourite calender application
 1. Optional: Place this script somewhere on a publicly-accessible webserver, and copy its URL into your calender application

Note: Since your password is here in cleartext, you should obviously only put this on web server you trust completely, and not one anyone else could access.

Notes
-----

Errors will be printed on the console and emailed to the address given as your
username.

The script will search for all events starting between six months prior to,
and one year later than, the current date.
