<?php

/**
 * Pull events from an Exchange Web Service and generate an iCal file for
 * import into another system, eg. Google Calendar.
 */

require __DIR__ . '/vendor/autoload.php';

use \jamesiarmes\PhpEws\Client;
use \jamesiarmes\PhpEws\Request\FindItemType;

use \jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfBaseFolderIdsType;

use \jamesiarmes\PhpEws\Enumeration\DefaultShapeNamesType;
use \jamesiarmes\PhpEws\Enumeration\DistinguishedFolderIdNameType;
use \jamesiarmes\PhpEws\Enumeration\ResponseClassType;

use \jamesiarmes\PhpEws\Type\CalendarViewType;
use \jamesiarmes\PhpEws\Type\DistinguishedFolderIdType;
use \jamesiarmes\PhpEws\Type\ItemResponseShapeType;

// Replace with the date range you want to search in.
$start_date = new DateTime();
$start_date->setTimestamp(time() - 26*7*24*60*60); // 6 months in the past
$end_date = new DateTime();
$end_date->setTimestamp(time() + 52*7*24*60*60); // 12 months in the future
$timezone = 'Eastern Standard Time';

// Set connection information.
$host = "outlook.office365.com";
$username = trim(file_get_contents(__DIR__ . '/ews2ical.username'));
$password = trim(file_get_contents(__DIR__ . '/ews2ical.secret'));
$errors_to = is_readable(__DIR__ . '/ews2ical.errors_to') ? file_get_contents(__DIR__ . '/ews2ical.errors_to') : $username;
$version = Client::VERSION_2016;

$client = new Client($host, $username, $password);
$client->setTimezone($timezone);

$request = new FindItemType();
$request->ParentFolderIds = new NonEmptyArrayOfBaseFolderIdsType();

// Return all event properties.
$request->ItemShape = new ItemResponseShapeType();
$request->ItemShape->BaseShape = DefaultShapeNamesType::ALL_PROPERTIES;

$folder_id = new DistinguishedFolderIdType();
$folder_id->Id = DistinguishedFolderIdNameType::CALENDAR;
$request->ParentFolderIds->DistinguishedFolderId[] = $folder_id;

$request->CalendarView = new CalendarViewType();
$request->CalendarView->StartDate = $start_date->format('c');
$request->CalendarView->EndDate = $end_date->format('c');

$response = $client->FindItem($request);

$vCalendar = new \Eluceo\iCal\Component\Calendar('-//Exchange Events//NONSGML Exchange Events//EN');
$vCalendar->setPublishedTTL('PT1H');
$vCalendar->setName('EWS Events');
$vCalendar->setDescription('EWS Events for ' . $username);

// Used only for DTSTAMP field
$tz  = 'America/Toronto';
$dtz = new \DateTimeZone($tz);
date_default_timezone_set($tz);

$log_errors = '';
$log_events = '';
$num_errors = 0;
$num_events = 0;

// Iterate over the results, printing any error messages or event ids.
$response_messages = $response->ResponseMessages->FindItemResponseMessage;
foreach ($response_messages as $response_message) {
    // Make sure the request succeeded.
    if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
        $code = $response_message->ResponseCode;
        $message = $response_message->MessageText;
        $error = "Failed to search for events with '$code: $message'\n";
        fwrite(STDERR, $error);
        $log_errors .= $error;
        $num_errors++;
        continue;
    }

    // Iterate over the events that were found, printing some data for each.
    $items = $response_message->RootFolder->Items->CalendarItem;
    foreach ($items as $item) {
        $id = $item->ItemId->Id;
        $start = new DateTime($item->Start);
        $end = new DateTime($item->End);
        $created = new DateTime($item->DateTimeCreated);
        $isCancelled = $item->IsCancelled ? "Yes" : "No";
        $log_events .= 'Found event ' . $item->ItemId->Id . "\n"
            . '  Change Key: ' . $item->ItemId->ChangeKey . "\n"
            . '  Location:  ' . $item->Location . "\n"
            . '  Title: ' . $item->Subject . "\n"
            . '  To: ' . $item->DisplayTo . "\n"
            . '  CC: ' . $item->DisplayCc . "\n"
            . '  Organizer: ' . $item->Organizer->Mailbox->Name . "\n"
            . '  Start:  ' . $start->format('l, F jS, Y g:ia') . "\n"
            . '  End:    ' . $end->format('l, F jS, Y g:ia') . "\n"
            . '  Cancelled:    ' . $isCancelled . "\n"
            . '  MyResponse:    ' . $item->MyResponseType . "\n"
            . "\n";
        $num_events++;

        $vEvent = new \Eluceo\iCal\Component\Event();
        $vEvent->setDtStart(new \DateTime($item->Start));
        $vEvent->setDtEnd(new \DateTime($item->End));
        $vEvent->setCreated(new \DateTime($item->DateTimeCreated));
        $vEvent->setSummary($item->Subject);
        $vEvent->setLocation($item->Location);
        if ($item->IsCancelled) {
          $vEvent->setStatus('CANCELLED');
        }
        elseif ($item->MyResponseType == 'Accept') {
          $vEvent->setStatus('CONFIRMED');
        }
        elseif ($item->MyResponseType == 'NoResponseReceived') {
          $vEvent->setStatus('TENTATIVE');
        }

        // Cleanup: EWS likes to add ", Mr." to people's names (and sometimes truncates that string)
        $vOrganizer = new \Eluceo\iCal\Property\Event\Organizer(str_replace(', Mr.', '', $item->Organizer->Mailbox->Name));
        $vEvent->setOrganizer($vOrganizer);

        $vAttendees = new \Eluceo\iCal\Property\Event\Attendees();
        $attendees = [];
        $attendees += explode('; ', $item->DisplayTo);
        $attendees += explode('; ', $item->DisplayCc);
        // Cleanup: EWS likes to add ", Mr." to people's names (and sometimes truncates that string)
        $attendees = str_replace(', Mr.', '', $attendees);
        $attendees = str_replace(', Mr', '', $attendees);
        $attendees = str_replace(', M', '', $attendees);
        $attendees = array_unique($attendees);
        foreach ($attendees as $attendee) {
          $vEvent->addAttendee($attendee);
        }

        $vCalendar->addComponent($vEvent);
    }
}

// Handle errors
if ($log_errors) {
  mail($errors_to, 'ews2ical: errors found', $log_errors);
  file_put_contents('ews2ical.errors', $log_errors);
}

// Print to standard output for the web application
$output = $vCalendar->render() . "\n";
header('Content-Type: text/calendar; charset=utf-8');
header('Content-Disposition: attachment; filename="cal.ics"');
print $output;

// Save a local copy
file_put_contents('ews2ical.ics', $output);

// Save a copy of all events
file_put_contents('ews2ical.events', $log_events);

// Save a log of this run
$agent = $_SERVER['HTTP_USER_AGENT'];
$remote_addr = $_SERVER['REMOTE_ADDR'];
$remote_host = gethostbyaddr($remote_addr);
$log = sprintf("%s: %s: %d events %d errors (%s) %s\n", date('r'), $remote_host, $num_events, $num_errors, $agent, $_SERVER['REQUEST_URI']);
file_put_contents('ews2ical.log', $log, FILE_APPEND);
