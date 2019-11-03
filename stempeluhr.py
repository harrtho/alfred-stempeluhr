#!/usr/bin/env python
# encoding: utf-8
#
# Copyright (c) 2019 xDevThomas@gmail.com.
#
# MIT Licence. See http://opensource.org/licenses/MIT
#
# Created on 2019-11-02
#

"""stempeluhr.py [command] [<query>]

Calculate overtime based on the backup file of StempelUhr.

Usage:
    stempeluhr.py
    stempeluhr.py year [<query>]
    stempeluhr.py <query>

Options:
    -h, --help      Show this message

"""

import sys
import os
import csv
import math
from datetime import datetime, timedelta

from workflow import Workflow3
from workflow.background import is_running, run_in_background

# GitHub repo for self-updating
UPDATE_SETTINGS = {'github_slug': 'xDevThomas/alfred-stempel-uhr'}

# GitHub Issues
HELP_URL = 'https://github.com/xdevthomas/alfred-stempel-uhr/issues'

# Icon shown if a newer version is available
ICON_UPDATE = 'update-available.png'

# The default monthly hours already monthly compensated
DEFAULT_MONTHLY_COMPENSATED = 0.0

# Index of start date
START = 0
# Index of end date
END = 1
# Index of location
LOCATION = 2
# Index of break
BREAK = 3
# Index of time
TIME = 4
# Index of comment
COMMENT = 5
# Index of absence name
ABSENCE_NAME = 6
# Index of absence start
ABSENCE_START = 7
# Index of absence end
ABSENCE_END = 8
# Index of required hours
REQUIRED_HOURS = 9

# Month from index in name
MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
]

class AttrDict(dict):
    """Access dictionary keys as attributes."""

    def __init__(self, *args, **kwargs):
        """Create new dictionary."""
        super(AttrDict, self).__init__(*args, **kwargs)
        # Assigning self to __dict__ turns keys into attributes
        self.__dict__ = self

def parse_stempel(opts):
    overview = {}
    overview["Start_Year"] = None
    overview["End_Year"] = None

    with open(opts.stempel_uhr, 'r') as csvfile:
        stamps = csv.reader(csvfile, delimiter=';')
        # Head line
        stamps.next()
        last_stamp = None
        for stamp in stamps:
            current_stamp = {}
            current_stamp["Start"] = datetime.strptime(stamp[START], '%Y-%m-%d %H:%M:%S')
            current_stamp["End"] = datetime.strptime(stamp[END], '%Y-%m-%d %H:%M:%S')
            current_stamp["Break"] = float(stamp[BREAK])
            current_stamp["Time"] = float(stamp[TIME])
            current_stamp["Requred_Hours"] = float(stamp[REQUIRED_HOURS])

            # Current stamp is newer than today
            if current_stamp["Start"] >= datetime.now():
                break

            if overview["Start_Year"] is None:
                overview["Start_Year"] = current_stamp["Start"].date().year

            elif overview["Start_Year"] > current_stamp["Start"].date().year:
                overview["Start_Year"] = current_stamp["Start"].date().year

            if overview["End_Year"] is None:
                overview["End_Year"] = current_stamp["Start"].date().year

            elif overview["End_Year"] < current_stamp["Start"].date().year:
                overview["End_Year"] = current_stamp["Start"].date().year

            # In case of first stamped day, calculate as day change
            if last_stamp is None:
                last_stamp = current_stamp
                current_stamp["Time"] -= current_stamp["Requred_Hours"]


            date_diff = current_stamp["Start"].date() - last_stamp["Start"].date()
            if date_diff.days > 0:
                current_stamp["Time"] -= current_stamp["Requred_Hours"]

            add_stamp(overview, current_stamp)

            last_stamp = current_stamp

    # log.debug(overview)
    return overview

def add_stamp(overview, stamp):
    year = stamp["Start"].date().year
    month = MONTHS[stamp["Start"].date().month - 1]

    if year not in overview:
        overview[year] = {}

    overview_year = overview[year]
    if month not in overview_year:
        overview_year[month] = {}

    overview_month = overview_year[month]
    if "Overtime" not in overview_month:
        overview_month["Overtime"] = 0.0

    overview_month["Overtime"] += stamp["Time"]

def calc_overtime(overview, opts):
    start_year = overview["Start_Year"]
    end_year = overview["End_Year"]
    overview["Total_Overtime"] = 0.0
    overview["Total_Lost"] = 0.0
    year = start_year


    while year <= end_year:
        current_year = overview[year]
        current_year["Total_Overtime"] = 0.0
        current_year["Total_Lost"] = 0.0

        for month in MONTHS:
            if month not in current_year:
                if (year <= datetime.now().year) and (MONTHS.index(month) + 1 <= datetime.now().month):
                    current_year[month] = {'Overtime' : 0.0, 'Lost' : 0.0}

                continue

            current_month = current_year[month]
            if overview["Total_Overtime"] >= 0.0:
                if current_month["Overtime"] >= opts.monthly_compensated:
                    current_month["Overtime"] -= opts.monthly_compensated
                    current_month["Lost"] = opts.monthly_compensated
                    overview["Total_Overtime"] += current_month["Overtime"]
                    current_year["Total_Overtime"] += current_month["Overtime"]
                    overview["Total_Lost"] += opts.monthly_compensated
                    current_year["Total_Lost"] += opts.monthly_compensated

                elif current_month["Overtime"] > 0.0:
                    overview["Total_Lost"] += current_month["Overtime"]
                    current_year["Total_Lost"] += current_month["Overtime"]
                    current_month["Lost"] = current_month["Overtime"]
                    current_month["Overtime"] = 0.0

                else:
                    overview["Total_Overtime"] += current_month["Overtime"]
                    current_year["Total_Overtime"] += current_month["Overtime"]
                    current_month["Lost"] = 0.0

            else:

                overview["Total_Overtime"] += current_month["Overtime"]
                current_year["Total_Overtime"] += current_month["Overtime"]
                current_month["Lost"] = 0.0
                if overview["Total_Overtime"] >= opts.monthly_compensated:
                    overview["Total_Overtime"] -= opts.monthly_compensated
                    current_year["Total_Overtime"] -= opts.monthly_compensated
                    overview["Total_Lost"] += opts.monthly_compensated
                    current_year["Total_Lost"] += opts.monthly_compensated
                    current_month["Lost"] = opts.monthly_compensated

                elif overview["Total_Overtime"] > 0.0:
                    overview["Total_Lost"] += overview["Total_Overtime"]
                    current_year["Total_Lost"] += overview["Total_Overtime"]
                    current_month["Lost"] = overview["Total_Overtime"]
                    current_year["Total_Overtime"] -= overview["Total_Overtime"]
                    overview["Total_Overtime"] = 0.0

        year +=1

def show_error(wf, opts):
    wf.add_item(u"Error no StempelUhrDataBackup.csv",
                u'Configure StempelUhrDataBackup.csv path in Workflow Settings',
                valid=False,
                icon='/System/Library/CoreServices/CoreTypes.bundle/Contents/Resources/AlertStopIcon.icns')
    return 0

def do_current(wf, opts):
    overview = parse_stempel(opts)
    calc_overtime(overview, opts)
    log.debug('overview=%r', overview)

    total_overtime_hours = int(math.floor(overview["Total_Overtime"]))
    total_overtime_minutes = int(math.floor((overview["Total_Overtime"] - math.floor(overview["Total_Overtime"]))*60))
    total_lost_hours = int(math.floor(overview["Total_Lost"]))
    total_lost_minutes = int(math.floor((overview["Total_Lost"] - math.floor(overview["Total_Lost"]))*60))

    wf.add_item(u'%sh:%smin' % (total_overtime_hours, total_overtime_minutes),
                u'Total Overtime',
                valid=False,
                icon='icon.png')
    wf.add_item(u'%sh:%smin' % (total_lost_hours, total_lost_minutes),
                u'Total Lost',
                valid=False,
                icon='icon.png')

    return 0

def do_year(wf, opts):
    overview = parse_stempel(opts)
    calc_overtime(overview, opts)
    log.debug('overview=%r', overview)

    try:
        year = int(opts.query)
        if int(opts.query) not in overview:
            add_year_options(overview, wf, opts)
        else:
            current_year = overview[int(opts.query)]
            for month in MONTHS[::-1]:
                if (datetime.now().year <= int(opts.query)) and \
                    (MONTHS.index(month) + 1 > datetime.now().month):
                    continue

                current_month = current_year[month]
                overtime_hours = int(math.floor(current_month["Overtime"]))
                overtime_minutes = int(math.floor((current_month["Overtime"] - math.floor(current_month["Overtime"]))*60))
                lost_hours = int(math.floor(current_month["Lost"]))
                lost_minutes = int(math.floor((current_month["Lost"] - math.floor(current_month["Lost"]))*60))
                wf.add_item(u'%s: %sh:%smin' % (month, overtime_hours, overtime_minutes),
                    u'Lost: %sh:%smin' % (lost_hours, lost_minutes),
                    valid=False,
                    icon='icon.png')
    except ValueError:
        add_year_options(overview, wf, opts)

def parse_args():
    """Extract options from CLI arguments.

    Returns:
        AttrDict: CLI options.

    """
    from docopt import docopt

    args = docopt(__doc__, wf.args)
    log.debug('args=%r', args)

    # The file path to the StempelUhrDataBackup.csv
    stempel_uhr = os.path.expanduser(os.getenv('FILE_PATH'))

    # The monthly hours aready moth compensated
    monthly_compensated = float(os.getenv('MONTHLY_COMPENSATED', DEFAULT_MONTHLY_COMPENSATED))

    opts = AttrDict(
        query=(args.get('<query>') or u'').strip(),
        stempel_uhr=stempel_uhr,
        monthly_compensated=monthly_compensated,
        do_year=args.get('year'),
    )

    log.debug('opts=%r', opts)

    return opts

def add_global_options(wf, opts):
    if 'year'.startswith(opts.query):
        wf.add_item(u'Year',
                u'Show result from caputured years...',
                valid=False,
                autocomplete='year ',
                icon='icon.png')

def add_year_options(overview, wf, opts):
    if opts.query == None:
        pass

    start_year = overview["Start_Year"]
    end_year = overview["End_Year"]
    year = end_year

    while year >= start_year:
        current_year = overview[year]
        overtime_hours = int(math.floor(current_year["Total_Overtime"]))
        overtime_minutes = int(math.floor((current_year["Total_Overtime"] - math.floor(current_year["Total_Overtime"]))*60))
        lost_hours = int(math.floor(current_year["Total_Lost"]))
        lost_minutes = int(math.floor((current_year["Total_Lost"] - math.floor(current_year["Total_Lost"]))*60))
        wf.add_item(u'%s total: %sh:%smin' % (year, overtime_hours, overtime_minutes),
            u'Lost: %sh:%smin' % (lost_hours, lost_minutes),
            valid=False,
            autocomplete="year %s" %year,
            icon='icon.png')
        year -= 1

def main(wf):
    return_value = 0

    # Parse args from workflow
    opts = parse_args()

    # Run workflow
    if not os.path.isfile(opts.stempel_uhr):
        return show_error(wf, opts)

    # Notify user if update is available
    # ------------------------------------------------------------------
    if wf.update_available:
        wf.add_item(u'Workflow Update is Available',
                    u'↩ or ⇥ to install',
                    autocomplete='workflow:update',
                    valid=False,
                    icon=ICON_UPDATE)

    if opts.do_year:
        do_year(wf, opts)

    else:
        do_current(wf, opts)
        add_global_options(wf, opts)

    wf.send_feedback()

    return return_value

if __name__ == '__main__':
    wf = Workflow3(help_url=HELP_URL,
                   update_settings=UPDATE_SETTINGS
                   )
    log = wf.logger
    sys.exit(wf.run(main))
