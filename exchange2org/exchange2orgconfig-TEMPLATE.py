# -*- coding: utf-8; mode: python; -*-
# Time-stamp: <2017-09-16 18:20:34 vk>

# ===================================================================== ##
#                                                                       ##
#  These are exchange2org.py configuration settings.                    ##
#                                                                       ##
#  You might not want to modify anything if you do not know, what       ##
#  you are doing :-)                                                    ##
#                                                                       ##
# ===================================================================== ##

# Server to connect to. Like: 'mail.example.com'
EXCHANGE_SERVER = 'mail.example.com'

# Domain and username like 'MYDOMAIN\\firstname.lastname'
USERNAME = 'MYDOMAIN\\firstname.lastname'
# In some cases (e.g. Office365) you only need your e-mail address. Try and
# see what works.
# USERNAME = 'first.last@example.com'

# You are about to enter your exchange password in clear text. Please
# make sure that only you are allowed to access this configuration
# file using proper permissions!
PASSWORD='PLEASE USE SECURE PASSPHRASES AND DONT SHARE THEM BETWEEN ACCOUNTS'

# Email address of the account
PRIMARY_SMTP_ADDRESS = 'first.last@example.com'

# Has to be an ISO time-zone like 'Europe/Vienna'
TIMEZONE = 'Europe/Vienna'

# The category which will be set (or an empty string for no category)
CATEGORY = 'mycompany'

# A list of tags which will be used (or an empty list)
TAGS = ['OUTLOOK']

# If you define a custom hyperlink for Outlook objects, you can click
# on links in order to jump to the Outlook element.
# Adding links: http://orgmode.org/manual/Adding-hyperlink-types.html
# Jumping to those links in Outlook: https://superuser.com/a/100084
# Set to empty string to disable those links.
OUTLOOK_HYPERLINK = 'outlook'

# A list of subjects whose entries are omitted in the output. I use it to
# suppress the output of holidays I got in Org-mode from a different source.
OMIT_SUBJECTS = ['Christmas Eve', 'Sylvester', 'Good Friday', 'Holy Thursday',
                 'St. Stephen\'s Day', 'Christmas Day', 'Labor Day', 'Whit Sunday',
                 'Whit Monday', 'New Year\'s Day', 'National Day',
                 'Immaculate Conception', 'Epiphany', 'Easter Monday',
                 'Easter Day', 'Corpus Christi', 'Assumption', 'Ascension',
                 'All Saints\' Day']

# If you don't need the PROPERTIES drawers containing the IDs of the
# entries, set this to False:
WRITE_PROPERTIES_DRAWER = True

# ===================================================================== ##
#                                                                       ##
#  These are INTERNAL configuration settings.                           ##
#                                                                       ##
#  You might not want to modify anything if you do not REALLY know,     ##
#  what you are doing :-)                                               ##
#                                                                       ##
# ===================================================================== ##


# the assert-statements are doing basic sanity checks on the configured variables
# please do NOT change them unless you are ABSOLUTELY sure what this means for the rest of lazyblorg!

assert type(EXCHANGE_SERVER) == str
assert type(USERNAME) == str
assert type(PASSWORD) == str
assert type(PRIMARY_SMTP_ADDRESS) == str
assert type(TIMEZONE) == str
assert type(CATEGORY) == str
assert type(TAGS) == list
assert type(OUTLOOK_HYPERLINK) == str
assert type(OMIT_SUBJECTS) == list
assert type(WRITE_PROPERTIES_DRAWER) == bool

# END OF FILE #################################################################
# Local Variables:
# mode: flyspell
# eval: (ispell-change-dictionary "en_US")
# End:
