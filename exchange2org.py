#!/usr/bin/env python3
# -*- coding: utf-8 -*-
PROG_VERSION = "Time-stamp: <2017-09-16 18:29:15 vk>"

# TODO:
# - fix parts marked with «FIXXME»
# - implement recurring events
# - write README.org


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
save_import('argparse')   # for handling command line arguments
save_import('time')
save_import('datetime')
save_import('logging')
save_import('exchangelib')  # for accessing Exchange servers


PROG_VERSION_DATE = PROG_VERSION[13:23]

DAY_STRING_REGES = re.compile('([12]\d\d\d)-([012345]\d)-([012345]\d)')

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
  " + sys.argv[0] + " --calendar some/subfolder/my_exchange_calendar.org\n\
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

    components = re.match(DAY_STRING_REGES, timestr)
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

    def convert_to_orgmode(self, event):

        subject = event.subject

        start_day = event.start.astimezone(self.tz).ewsformat()[:10]
        start_time = event.start.astimezone(self.tz).ewsformat()[11:16]
        end_day = event.end.astimezone(self.tz).ewsformat()[:10]
        end_time = event.end.astimezone(self.tz).ewsformat()[11:16]

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
        debugtext.append('ID: ' + event.item_id)
        debugtext.append('is_all_day: ' + repr(event.is_all_day)) # =False
        if event.location:
            debugtext.append('location: ' + event.location)  #=None,
        else:
            debugtext.append('(no location)')
        debugtext.append('is_cancelled: ' + repr(event.is_cancelled))  #=False,

        self.logger.debug('=' * 80 + '\n' + '\n'.join(debugtext))

        if event.is_cancelled or event.subject in self.config.OMIT_SUBJECTS:
            return False

        # legacy_free_busy_status='Busy',
        # organizer=Mailbox('Karl Voit', 'K.Voit@detego.com', 'Mailbox', None),
        # required_attendees=[Attendee(Mailbox('Karl Voit', 'K.Voit@detego.com', 'Mailbox', None), 'Unknown', None)],
        # optional_attendees=None, resources=None, recurrence=Recurrence(WeeklyPattern(1,  [1,  2,  3,  4, 5],  7),  NoEndPattern(EWSDate(2017,  9,  13),)),

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
            output += ' [[outlook:' + event.item_id + '][🔗]]'

        if self.config.WRITE_PROPERTIES_DRAWER:
            output += '\n:PROPERTIES:\n:ID: ' + event.item_id + '\n:END:\n'
        else:
            output += '\n'

        if options.verbose:
            output += ': ' + '\n: '.join(debugtext) + '\n'

        return output

    def dump_calendar(self):

        number_of_events = 0
        outputfilename = options.outputfile[0]

        # Fetch all calendar events from the Exchange server:
        events = self.account.calendar.view(start=self.tz.localize(exchangelib.EWSDateTime(2017, 1, 1)), end=self.tz.localize(exchangelib.EWSDateTime(2018, 1, 1)))

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

        self.logger.info(str(number_of_events) + ' events were written to ' + outputfilename)

def main():
    """Main function"""

    handle_logging()

    if options.verbose and options.quiet:
        error_exit(1, "Options \"--verbose\" and \"--quiet\" found. " +
                   "This does not make any sense, you silly fool :-)")

    if options.dryrun:
        logging.debug("DRYRUN active, not changing any files")

    # Looking for the configuration file which is on a hard-coded path:
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
        exchange2org.dump_calendar()

    if not options.quiet:
        # add empty line for better screen output readability
        print()

    if True:
        logging.debug('successfully finished.')
    else:
        logging.debug("finished with FIXXME")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:

        logging.info("Received KeyboardInterrupt")

# END OF FILE #################################################################
