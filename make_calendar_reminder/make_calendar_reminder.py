#!/usr/bin/python
#
# $Revision:: 119                                          $
# $Date:: 2014-08-19 17:12:58 -0700 (Tue, 19 Aug 2014)     $
# $Author:: wkumari                                        $
# $HeadURL:: svn+ssh://svn.kumari.net/data/svn/code/python#$
# Copyright: Warren Kumari (warren@kumari.net) -- 2013
#

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from icalendar import Calendar, Event, vCalAddress, vText
from optparse import OptionParser
import ConfigParser
import base64
import datetime
import email
import os
import pytz
import smtplib
import sys
import uuid


# Example config file
EXAMPLE_CFG="""
[Server]
use_ssl = True
server = smtp.gmail.com

[User]
username = fred@example.com
from = Freds Automated Scripts <fred@example.com>
password = Hunter2
"""

SMTP_DEBUG = False


def ParseOptions():
  """Parses the command line options."""
  global opts
  usage =  """%prog -d days -m="message" <-r>
  This program creates an iCal format reminder.

  Example:
      %prog -d 14 -m "WGLC finishes for draft-opsec-foo-bar-01" -r
         Creates an iCal reminder in 14 days to end WGLC.
         Also adds a reminder in 7 days.
  """

  options = OptionParser(usage=usage)
  options.add_option('-r', '--reminder', dest='reminder',
                     action='store_true',
                     default=False,
                     help="""Create a reminder between now and the event.
                     NOTE: gMail does not deal with multiple events in an iCal attachment...""")
  
  options.add_option('-d', '--days', dest='days',
                     action='store', type='int',
                     default='14',
                     help='In how many days to create the reminder (REQUIRED)')

  options.add_option('-m', '--message', dest='message',
                     default='',
                     help='Text for the reminder (REQUIRED)')

  options.add_option('-t', '--to', dest='to',
                     default='',
                     help=('Email address to send the reminder to (REQUIRED)'))

  options.add_option('-c', '--config', dest='config',
                    default='calendar.cfg',
                    help=('INI format config file.'))
  
  (opts, args) = options.parse_args()
  if not (opts.message and opts.days and opts.to):
    print ('Need days and message')
    options.print_help()
    sys.exit(-1)
  opts.files = args
  return opts


def ParseConfig(filename):
  """Reads in the INI format config file."""
  
  cfg = {}
  error=False
  config = ConfigParser.RawConfigParser()
  config.read(filename)
  if not config:
    error=True
  else:
    try:
      cfg['use_ssl'] = config.getboolean ('Server', 'use_ssl')
      cfg['server'] = config.get ('Server', 'server')

      cfg['user'] = config.get ('User', 'username')
      cfg['from'] = config.get ('User', 'from')
      cfg['password'] = config.get ('User', 'password')
    except (ConfigParser.NoSectionError, ConfigParser.NoOptionError), e:
      print 'Unable to parse config file: %s' % e
      error=True
    
  if not cfg or error:
    print 'Unable to read config file %s.\n\nExample config:\n%s' % (
      filename, EXAMPLE_CFG)
    sys.exit(1)
  return cfg


def CreateCalendar():
  """Returns a Calendar, suitable for adding Reminders to."""
  cal = Calendar()
  cal.add('prodid', '-/IETF Calendar Reminder//kumari.net//')
  cal.add('version', '2.0') 
  return cal

def CreateReminder(cal, days, text, fromaddress):
  """Creates a calendar entry in days days with title text.

  Args:
    cal: A Calendar.
    days: How many days in the future to create the reminder.
    text: The calendar summary

  Returns:
    A calendar entry.
    """
  event = Event()
  event.add('summary', text)
  eventdate = datetime.datetime.now() + datetime.timedelta(days=int(days))
  event.add('dtstart',eventdate)
  event.add('dtend', eventdate)
  event.add('dtstamp', eventdate)
  event['uid'] = str(uuid.uuid1()) + '@kumari.net'
  event.add('priority', 5)

  organizer = vCalAddress('MAILTO:%s' % fromaddress)
  organizer.params['cn'] = vText('%s' % fromaddress)
  organizer.params['role'] = vText('CHAIR')
  event['organizer'] = organizer
  event['location'] = vText('Online')

  cal.add_component(event)
  return cal


def WriteICSFile(cal): 
  f = open('example.ics', 'w.b.')
  f.write(cal.to_ical())
  f.close()

  
def CreateEvents(days, message, fromaddress):
  """Creates the entries"""
  cal = CreateCalendar()
  cal = CreateReminder(cal, days, message, fromaddress)
  return cal

def SendiCalEmail(cal, username, sender, to, server, ssl, password):
  """Sends the iCal calendar to to, from from, using server server!
                    
  Args:
    cal: The calandar
    username: User to send mail from.
    sender: String to send the mail from.
    to: An array containing who to send the mail to.
    server: SMTP server to use.
    ssl: Boolean, use SSL or not."""


  # Parse addresses of the form "foo@exmaple.com, bar@baz.com" into a list.
  to_list = to.split(',')
  
  # Bah, we need a multipart message to attach the iCal entry to.
  msg = MIMEMultipart()
  msg['Subject'] = 'Calandar reminder: %s' % opts.message
  msg['From'] = sender
  msg['To'] = to
  msg.preamble = '''Calendar invitation for reminder: %s.
     Created by script: %s''' % (opts.message, sys.argv[0])

  # We need to create the attachment.
  txt = MIMEText(msg.preamble)
  msg.attach(txt) 
  ical = MIMEText(cal.to_ical(), 'calendar', 'utf-8')
  msg.attach(ical)

  if ssl:
    s=smtplib.SMTP_SSL(server)
    s.login(username, password)
  else:
    s = smtplib.SMTP(server, 25)
  s.set_debuglevel(SMTP_DEBUG)
  s.sendmail(sender, to_list, msg.as_string())
  s.quit()


if __name__ == "__main__":
  ParseOptions()
  cfg = ParseConfig(os.path.join(os.path.dirname(os.path.realpath(__file__)), opts.config))
  
  cal=CreateEvents(opts.days, opts.message, cfg['from'])
  SendiCalEmail(cal, cfg['user'], cfg['from'], opts.to, cfg['server'],
                cfg['use_ssl'], cfg['password'])
  if opts.reminder:
    cal=CreateEvents(opts.days/2, "Reminder: " + opts.message, cfg['from'])
    SendiCalEmail(cal, cfg['user'], cfg['from'], opts.to, cfg['server'],
                  cfg['use_ssl'], cfg['password'])

