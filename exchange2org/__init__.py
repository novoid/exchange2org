#!/usr/bin/env python3
# -*- coding: utf-8 -*-
PROG_VERSION = "Time-stamp: <2018-02-03 19:41:17 vk>"

# TODO:
# - fix parts marked with «FIXXME»


# ===================================================================== ##
#  You might not want to modify anything below this line if you do not  ##
#  know, what you are doing :-)                                         ##
# ===================================================================== ##

import importlib

def save_import(library):
    try:
        globals()[library] = importlib.import_module(library)
    except ImportError:
        print("Could not find Python module \"" + library + "\".\nPlease install it, e.g., with \"sudo pip install " + library + "\".")
        sys.exit(2)


import sys
import os
import re
save_import('base64')       # itemID/entryID conversion
save_import('argparse')     # for handling command line arguments
save_import('time')
save_import('datetime')
save_import('logging')
save_import('exchangelib')  # for accessing Exchange servers


PROG_VERSION_DATE = PROG_VERSION[13:23]

DAY_STRING_REGEX = re.compile('([12]\d\d\d)-([012345]\d)-([012345]\d)')

DESCRIPTION = "This tool connects to your Exchange server and extracts data\n\
in Org-mode format.\n\
\n\
Please do note that this is a TEMPORARY stand-alone tool which will\n\
be added to Memacs as an module as soon as Memacs got migrated to\n\
Python 3:   https://github.com/novoid/memacs/\n\
\n\
You can configure the behavior and output via a configuration file.\n\
\n\
Example usages:\n\
  exchange2org --calendar some/subfolder/my_exchange_calendar.org\n\
      … writes your calendar events into the Org-mode file.\n\
\n\
\n"

EPILOG = u"\n\
:copyright: (c) by Karl Voit <tools@Karl-Voit.at>\n\
:license: GPL v3 or any later version\n\
:URL: https://github.com/novoid/exchange2org\n\
:bugreports: via github or <tools@Karl-Voit.at>\n\
:version: " + PROG_VERSION_DATE + "\n·\n"

parser = argparse.ArgumentParser(prog=sys.argv[0],
                                 # keep line breaks in EPILOG and such
                                 formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog=EPILOG,
                                 description=DESCRIPTION)

parser.add_argument(dest='outputfile', metavar='FILE', nargs=1, help='The filename of the output file')

parser.add_argument('--calendar', action='store_true', help='Extract the calendar as Org-mode events. ')

parser.add_argument('--startday', metavar='date-or-days', nargs=1, help='Starting date for fetching data. ' +
                    'Default: 60 days in past. "date-or-days" is either of form "YYYY-MM-DD" or a number.')

parser.add_argument('--endday', metavar='date-or-days', nargs=1, help='End date for fetching data. ' +
                    'Default: 60 days in future. "date-or-days" is either of form "YYYY-MM-DD" or a number.')

parser.add_argument('--ignore-category', metavar='CATEGORY', nargs=1, help='Category whose events will be omitted.')

parser.add_argument('-s', '--dryrun', dest='dryrun', action='store_true',
                    help='enable dryrun mode: simulate what would happen, do not modify anything')

parser.add_argument('-v', '--verbose', dest='verbose', action='store_true',
                    help='enable verbose mode')

parser.add_argument('-q', '--quiet', dest='quiet', action='store_true',
                    help='enable quiet mode')

options = parser.parse_args()



def handle_logging():
    """Log handling and configuration"""

    if options.verbose:
        FORMAT = "%(levelname)-8s %(asctime)-15s %(message)s"
        logging.basicConfig(level=logging.DEBUG, format=FORMAT)
    elif options.quiet:
        FORMAT = "%(levelname)-8s %(message)s"
        logging.basicConfig(level=logging.ERROR, format=FORMAT)
    else:
        FORMAT = "%(levelname)-8s %(message)s"
        logging.basicConfig(level=logging.INFO, format=FORMAT)


def error_exit(errorcode, text):
    """exits with return value of errorcode and prints to stderr"""

    sys.stdout.flush()
    logging.error(text)

    sys.exit(errorcode)


def day_string_to_datetime(timestr):
    """
    Returns a datetime object containing the time-stamp of a day in a string.

    @param orgtime: YYYY-MM-DD
    @param return: date time object
    """

    components = re.match(DAY_STRING_REGEX, timestr)
    # components.groups() -> ('2017', '08', '15')

    if not components:
        logging.error(
            "string could not be parsed as time-stamp of format \"YYYY-MM-DD\": \"%s\"",
            timestr)

    try:
        year = int(components.group(1))
        month = int(components.group(2))
        day = int(components.group(3))
    except:
        logging.error(
            "strings could not be casted to int from format \"YYYY-MM-DD\": \"%s\"",
            timestr)

    return datetime.datetime(year, month, day)


class Exchange2Org(object):

    logger = None
    config = None

    tz = None
    exchange_config = None

    def __init__(self, configuration, logger):
        self.logger = logger
        self.config = configuration

        try:
            self.exchange_config = exchangelib.Configuration(exchangelib.Credentials(self.config.USERNAME, self.config.PASSWORD), server=self.config.EXCHANGE_SERVER)
            self.account = exchangelib.Account(self.config.PRIMARY_SMTP_ADDRESS, config=self.exchange_config, autodiscover=False, access_type=exchangelib.DELEGATE)
            self.tz = exchangelib.EWSTimeZone.timezone(self.config.TIMEZONE)
        except:
            logger.critical('Error occured while trying to set up connection with the exchange server "' + self.config.EXCHANGE_SERVER + '":')
            raise

    def convert_itemid_from_exchange_to_entryid_for_outlook(self, itemid):
        """
        Converts the string of the ItemID we got from the exchange server to the
        EntryID we can query in Outlook.

        Don't ask me why there is a need for the two IDs. It took me some hours to
        figure out the issue in the first place and another two hours to find the
        magic algorythm to convert.

        Interesting sources:
        https://blogs.msdn.microsoft.com/brijs/2010/09/09/how-to-convert-exchange-items-entryid-to-ews-unique-itemid-via-ews-managed-api-convertid-call/
        Solution: https://github.com/ecederstrand/exchangelib/issues/146

        @param itemid: string of the itemid from the exchange server
        @param return: string of the entryid for outlook
        """

        # Use the magic wand:
        decoded_val = base64.b64decode(itemid)
        itemID = decoded_val.hex().upper()

        # Somehow, my Outlook does not want the first 86 characters:
        return itemID[86:]

    def convert_to_orgmode(self, event):
        """
        Gets a calendar event and returns its representation in Org-mode format.

        @param event: an Exchange calendar event
        @param return: string containing a heading with a representation of the calendar event
        """

        subject = event.subject

        if event.is_cancelled or \
           event.subject in self.config.OMIT_SUBJECTS:
            return False

        if event.categories and options.ignore_category:
            if options.ignore_category[0] in event.categories:
                return False

        start_day = event.start.astimezone(self.tz).ewsformat()[:10]
        start_time = event.start.astimezone(self.tz).ewsformat()[11:16]
        end_day = event.end.astimezone(self.tz).ewsformat()[:10]
        end_time = event.end.astimezone(self.tz).ewsformat()[11:16]
        entry_id = self.convert_itemid_from_exchange_to_entryid_for_outlook(str(event.item_id))
        #entry_id = event.item_id  # until I found a working version for Python 3 of the function above

        if event.is_all_day:
            assert(end_time == '00:00')
            # When is_all_day is true, the end day is midnight after
            # the end day. I want to correct this to the end day which
            # is affected of the event instead:
            end_day_datetime = day_string_to_datetime(end_day)
            end_day_datetime = end_day_datetime + datetime.timedelta(days=-1)
            new_end_day = end_day_datetime.strftime('%Y-%m-%d')
            self.logger.debug('is_all_day: I moved end_day from ' + end_day + ' to ' + new_end_day)
            end_day = new_end_day

        debugtext = []
        debugtext.append('start:' + event.start.astimezone(self.tz).ewsformat())  # example: '2017-09-13T09:30:00+02:00'
        debugtext.append('start_day:' + start_day)
        debugtext.append('start_time:' + start_time)
        debugtext.append('end: ' + event.end.astimezone(self.tz).ewsformat())
        debugtext.append('end_day:' + end_day)
        debugtext.append('end_time:' + end_time)
        debugtext.append('subject: ' + event.subject)
        debugtext.append('item_id: ' + event.item_id)
        debugtext.append('entry_id: ' + entry_id)
        debugtext.append('is_all_day: ' + repr(event.is_all_day)) # =False
        if event.location:
            debugtext.append('location: ' + event.location)  #=None,
        else:
            debugtext.append('(no location)')
        debugtext.append('is_cancelled: ' + repr(event.is_cancelled))  #=False,

        self.logger.debug('=' * 80 + '\n' + '\n'.join(debugtext))

        output = '** <' + start_day

        if event.is_all_day:
            if start_day == end_day:
                output += '>'
            elif start_day != end_day:
                output += '>-<' + end_day + '>'
            else:
                error_exit(10, 'convert_to_orgmode found a day/time combination which is not implemented yet (is_all_day)')
        elif start_day == end_day:
            if start_time == end_time:
                output += ' ' + start_time + '>'
            else:
                output += ' ' + start_time + '-' + end_time + '>'
        elif start_day != end_day:
            output += ' ' + start_time + '>-<' + end_day + ' ' + end_time + '>'
        else:
            error_exit(11, 'convert_to_orgmode found a day/time combination which is not implemented yet')

        if subject:
            output += ' ' + subject

        if event.location:
            output += ' (' + event.location + ')'

        if len(self.config.OUTLOOK_HYPERLINK) > 1:
            output += ' [[outlook:' + entry_id + '][⦿]]'

        if self.config.WRITE_PROPERTIES_DRAWER:
            output += '\n:PROPERTIES:\n:ID: ' + entry_id + '\n:END:\n'
        else:
            output += '\n'

        if options.verbose:
            output += ': ' + '\n: '.join(debugtext) + '\n'

        return output

    def dump_calendar(self, startday, endday):
        """
        Retrieves Exchange calendar data from the server and writes output file.

        @param startday: list of year, month, day as integers
        @param endday: list of year, month, day as integers
        """

        number_of_events = 0
        outputfilename = options.outputfile[0]

        # Fetch all calendar events from the Exchange server:
        events = self.account.calendar.view(start=self.tz.localize(exchangelib.EWSDateTime(*startday)), end=self.tz.localize(exchangelib.EWSDateTime(*endday)))

        with open(outputfilename, 'w') as outputhandle:

            # Write Org-mode header with optional tags and category:
            if not options.dryrun:
                outputhandle.write('# -*- mode: org; coding: utf-8; -*-\n* Calendar events of "' +
                                   self.config.USERNAME.replace('\\', '\\\\') + '" from "' + self.config.EXCHANGE_SERVER +
                                   '"  ·•·  Generated via ' + sys.argv[0] + ' at ' +
                                   datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S'))
                if len(self.config.TAGS) > 0:
                    outputhandle.write(' ' * 10 + ':' + ':'.join(self.config.TAGS) + ':')
                if len(self.config.CATEGORY) > 0:
                    outputhandle.write('\n:PROPERTIES:\n:CATEGORY: ' + self.config.CATEGORY + '\n:END:\n')
                else:
                    outputhandle.write('\n')

            # Loop over all events from the Exchange server:
            for event in events:
                output = self.convert_to_orgmode(event)
                if output:
                    number_of_events += 1
                    if not options.dryrun:
                        outputhandle.write(output)

            if not options.dryrun:
                outputhandle.write('\n\n# Local Variables:\n# mode: auto-revert-mode\n# End:\n')

        self.logger.info(str(number_of_events) + ' events were written to ' + outputfilename)


def handle_date_or_period_argument(daystring, future):
    """
    Gets a string containing either a 'YYYY-MM-DD' ISO day format
    or an integer. Returns a list of year, month, day of the ISO day
    or (in case of an integer) a similar list with number of days in
    the past or the future - according to the 2nd parameter.

    @param daystring: string with either 'YYYY-MM-DD' format or an integer
    @param furure: boolean
    @param endday: list of year, month, day as integers
    """

    # check format: integer or ISO date string?
    components = re.match(DAY_STRING_REGEX, daystring)
    if components:
        # parameter was of format YYYY-MM-DD
        return [int(components.group(1)), int(components.group(2)), int(components.group(3))]
    else:
        # parameter might be an integer:
        try:
            number_of_days = int(daystring)
        except:
            error_exit(2, 'Format of daystring parameter must be "YYYY-MM-DD" or an integer.')
        if future:
            delta_days = number_of_days
        else:
            delta_days = -number_of_days
        daystring_datetime = datetime.datetime.now() + datetime.timedelta(days=delta_days)
        return [daystring_datetime.year, daystring_datetime.month, daystring_datetime.day]


def main():
    """Main function"""

    handle_logging()

    if options.verbose and options.quiet:
        error_exit(1, "Options \"--verbose\" and \"--quiet\" found. " +
                   "This does not make any sense, you silly fool :-)")

    # The defaults are: ±60 days
    before_60_days = datetime.datetime.now() + datetime.timedelta(days=-60)
    startday = [before_60_days.year, before_60_days.month, before_60_days.day]
    in_60_days = datetime.datetime.now() + datetime.timedelta(days=60)
    endday = [in_60_days.year, in_60_days.month, in_60_days.day]

    if options.startday:
        startday = handle_date_or_period_argument(options.startday[0], future=False)
        logging.debug('options.startday found and set to ' + repr(startday))
    if options.endday:
        endday = handle_date_or_period_argument(options.endday[0], future=True)
        logging.debug('options.endday found and set to ' + repr(endday))

    if options.dryrun:
        logging.debug("DRYRUN active, not changing any files")

    logging.debug('Looking for the configuration file which is expected to be found on a hard-coded path')
    CONFIGDIR = os.path.join(os.path.expanduser("~"), ".config/exchange2org")
    sys.path.insert(0, CONFIGDIR)  # add CONFIGDIR to Python path in order to find config file
    try:
        import exchange2orgconfig
    except ImportError:
        print('Could not find file  "' + os.path.join(CONFIGDIR, 'exchange2orgconfig.py') + '".' +
              '\nPlease take a look at "exchange2orgconfig-TEMPLATE.py", copy it, and configure accordingly.')
        sys.exit(1)

    # So far, we only handle calendar events. Maybe the future will bring more:
    if options.calendar:
        exchange2org = Exchange2Org(exchange2orgconfig, logging.getLogger())
        exchange2org.dump_calendar(startday=startday, endday=endday)
    else:
        logging.info('Sorry, at the moment this tool only supports calendar events. So please use the --calendar parameter.')

    if not options.quiet:
        # add empty line for better screen output readability
        print()

    logging.debug('successfully finished.')

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:

        logging.info("Received KeyboardInterrupt")

# END OF FILE #################################################################
