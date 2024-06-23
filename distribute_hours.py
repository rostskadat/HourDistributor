#!/usr/bin/env python
"""
Generate a justification sheet for the given employee and the given time period
"""

try:
    from datetime import date, datetime, timedelta
    from dotenv import load_dotenv
    import calendar
    import coloredlogs
    import logging
    import math
    import os
    import sys
    import yaml
    from argparse import ArgumentParser, RawTextHelpFormatter
    from pprint import pprint, pformat
    from workalendar.europe import Aragon
    import openpyxl
    from openpyxl.styles import NamedStyle, Font, Border, Side
except ModuleNotFoundError as e:
    print(f"{e}. Did you load your environment?")
    sys.exit(1)

load_dotenv()
logger = logging.getLogger()

DEFAULT_FMT = '%Y-%m-%d'
CELL_FMT = "0.0" # "[h]:mm"
PREFERRED_FMT = 'decimal'

class Project:
    """Base class to all projects"""

    def __init__(self, name:str, start:date, end:date):
        self.name = name
        self.start = start
        self.end = end
        self.daily_average = 0
        self.interval_limit = 0
        self.dedicated = 0
        self.projected = 0

    def __str__(self) -> str:
        return self.__repr__()
    
    def __repr__(self) -> str:
        # from pprint import pformat
        # pformat(vars(self), indent=2, width=1)
        return f"<Project '{self.name}' ({self.start.strftime(DEFAULT_FMT)} -> {self.end.strftime(DEFAULT_FMT)}, davg={"{:.2f}".format(self.daily_average)}, ilim={"{:.2f}".format(self.interval_limit)}, projected={self.projected})>"
    
    def is_open(self, working_day:date) -> bool:
        """Returns whether a working day falls within a project

        Args:
            working_day (date): the day to test

        Returns:
            bool: True if the working day falls within a project
        """
        return self.start <= working_day and working_day <= self.end
    
    def is_full(self) -> bool:
        """Returns whether the project has used up all its projected time

        Returns:
            bool: True if no more time cn be dedicated to the project
        """
        return ((self.dedicated >= self.projected) or (self.dedicated >= self.interval_limit))
    
    def get_max_dedication(self, available_today:int, granularity:int) -> int:
        """Returns the maximum number of minute that can be dedicated by the project taking into account different limits.

        Args:
            limit_daily (int): the maximum number of minutes that can be daily dedicated to all projects
            granularity (int): the minimum interval for justification

        Case: return project.available 
        +--------------> interval_limit
        +------------>   available_today
        +         +      granular_span
        +--->            project.available

        Case: return project.available 
        +--------------> interval_limit
        +----->          available_today
        +         +      granular_span
        +--->            project.available

        Case: return available_today
        +--------------> interval_limit
        +----->          available_today
        +         +      granular_span
        +----------->    project.available

        Case: return granular_span
        +--------------> interval_limit
        +------------>   available_today
        +         +      granular_span
        +----------->    project.available

        Case: return granular_span
        +--------------> interval_limit
        +------------>   available_today
        +         +      granular_span
        +----------->    project.available

        Case: return interval_limit
        +------->        interval_limit
        +------------>   available_today
        +         +      granular_span
        +----------->    project.available
        
        Returns:
            int: the number of minutes
        """
        available_this_year = self.projected - self.dedicated
        available_this_month = self.interval_limit - self.dedicated
        granular_span = max(math.ceil(min([available_this_year, available_this_month, available_today]) / granularity), 1) * granularity
        return min([available_this_year, available_this_month, available_today, granular_span])


def _load_project_breakdown(project_breakdown:str, employee:str, only_projects:str, justification_year:int, working_days:list, defaults) -> dict:
    """Returns a dict containing information about the projects for the given employee

    Args:
        project_breakdown (str): the filename of the Excel file containing the project breakdown
        employee (str): the name of the employee
        only_projects (str): a comma separated list of project name to filter on.
        defaults (dict): the default

    Raises:
        ValueError: raised if the employee does not have a project breakdown

    Returns:
        A dict containg the projects details (name, start, end):
        { project_name -> Project (name, start, end, daily_average, interval_limit, dedicated, projected)

    """
    logger.info(f"Reading project breakdown for '{employee}' ...")
    workbook = openpyxl.load_workbook(project_breakdown)
    if employee not in workbook:
        raise ValueError(f"Invalid employee parameter: No worksheet '{employee}' found in {project_breakdown}.")
    worksheet = workbook[employee]
    # determining the column indexes for each relevant column
    project_name_column = 0
    project_start_column = 0
    project_end_column = 0
    year_columns = {}
    # I determine the column indices...
    for col in worksheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            if cell.value == defaults['project-name-column']:
                project_name_column = cell.column
            elif cell.value == defaults['project-start-column']:
                project_start_column = cell.column
            elif cell.value == defaults['project-end-column']:
                project_end_column = cell.column
            elif (isinstance(cell.value, int) or isinstance(cell.value, str)) and int(cell.value) > 2000:
                year_columns[int(cell.value)] = cell.column
            else:
                logger.warning(f"Ignoring unknown header {cell.column} (='{cell.value}' /{type(cell.value)})")
    # The map { project_name -> Project }
    projects = { }
    only_projects = only_projects.split(",") if only_projects else None
    for i in range(worksheet.min_row+1, worksheet.max_row+1):
        project_name = worksheet.cell(row=i, column=project_name_column).value
        if only_projects and project_name not in only_projects:
            logger.info(f"Skipping non included project '{project_name}'")
            continue
        project_start = worksheet.cell(row=i, column=project_start_column).value
        project_end = worksheet.cell(row=i, column=project_end_column).value
        project = Project(project_name, project_start, project_end)
        projected_dedication = worksheet.cell(row=i, column=year_columns[justification_year]).value
        if not projected_dedication:
            projected_dedication = 0
        if projected_dedication > 0:
            project.projected = round(projected_dedication * 60)
            project_working_days = list(filter(lambda d: project.start <= d and d <= project.end, working_days))
            assert len(project_working_days) > 0, f"Project '{project_name}' has projected dedication ({int(project.projected)}'), but no working days ({project.start.strftime(DEFAULT_FMT)} -> {project.end.strftime(DEFAULT_FMT)})"
            project.daily_average = float(project.projected) / len(project_working_days)
            project.interval_limit = float(project.projected) / (project_working_days[-1].month - project_working_days[0].month + 1)
        projects[project_name] = project            

    return projects


def _get_justifiation_dates(interval, period):
    """Return the justification start and end date

    Args:
        interval (str): a string indicating whether the justification should be a month or a week
        period (date): the start date of the justification period

    Returns:
        date: the start date for the justification period
        date: the end date for the justification period
    """
    if interval == 'monthly':
        start = period.replace(day=1)
        end = period.replace(day=calendar.monthrange(start.year, start.month)[1])
    else:
        start = period - timedelta(days=period.weekday())
        end = period + timedelta(days=6-period.weekday()) 
    return start, end


def _get_working_days(justification_start, extra_holidays):
    # ok, I always justify a full year and then extract the intersting period.
    # This is due to the fact that we do not want to justify periods that are not a multiple of the 
    # justification granulariry. Only during the last period of the year, is this allowed.
    logger.debug(f"Creating the holiday calendar for Aragon in {justification_start.year} ...")
    work_calendar = Aragon()
    # With the holiday calendar I create the list of the opened days in the justification year
    extra_holidays = extra_holidays.split(",")
    start_of_the_year = justification_start.replace(month=1, day=1)

    working_days = []
    for delta_day in (range(1, 367 if calendar.isleap(justification_start.year) else 366)):
        target_day = start_of_the_year + timedelta(days=delta_day)
        if work_calendar.is_working_day(target_day) and target_day.strftime(DEFAULT_FMT) not in extra_holidays:
            working_days.append(target_day)
    return working_days, work_calendar


def _set_cell_value(cell, minutes:float):
    if PREFERRED_FMT == 'decimal':
        cell.value = minutes/60 
        cell.number_format = "0.0"
    else:
        # Excel processes time entries as a decimal fraction of a day
        cell.value = minutes/(60 * 24)
        cell.number_format = "[h]:mm"


def _generate_report(employee, overwrite, project_breakdown, projects, daily_dedications, justification_year, defaults):
    logger.info("Generating report ...")
    workbook = openpyxl.load_workbook(project_breakdown, keep_vba=True)
    if defaults['worksheet-report-template'] not in workbook:
        raise ValueError(f"Invalid configuration: No worksheet '{defaults['worksheet-report-template']}' found in {project_breakdown}.")
    report_template = workbook[defaults['worksheet-report-template']]
    
    report_title = f"{employee} - {justification_year}"
    if report_title in workbook:
        if not overwrite:
            logger.info(f"Report '{report_title}' already exists in {project_breakdown}. Use '--overwrite' to overwrite.")
            return
        report = workbook[report_title]
    else:
        report = workbook.copy_worksheet(report_template)
        report.title = report_title

    # Should try to use defined name instead...
    # definition = report.defined_names["employee"]
    report['B1'].value = employee
    report['B2'].value = employee
    report['B3'].value = justification_year
    report['B4'].value = defaults['max-yearly-limit']

    # I massage the daily_dedications by project/month

    report_dedications = {}
    for working_day, project_dedications in daily_dedications.items():
        for project_name, time_to_dedicate in project_dedications.items():
            if project_name not in report_dedications.keys():
                report_dedications[project_name] = {}
            if working_day.month not in report_dedications[project_name]:
                report_dedications[project_name][working_day.month] = 0
            report_dedications[project_name][working_day.month] += time_to_dedicate

    logger.debug(f"report_dedications={pformat(report_dedications)}")
    PRJ_BRK_ROW = 8
    report.delete_rows(PRJ_BRK_ROW)
    for project_name in sorted(report_dedications.keys(), reverse=True):
        report.insert_rows(PRJ_BRK_ROW)
        # let's write the project name
        project_cell = report[f"A{PRJ_BRK_ROW}"]
        project_cell.value = project_name
        try:
            project_cell.style = defaults['project-report-project-style']
        except:
            logger.warning(f"Missing named style '{defaults['project-report-project-style']}' in Excel file {project_breakdown}")

        # let's write the 'TOTAL HOURS' for that project
        _set_cell_value(report[f"D{PRJ_BRK_ROW}"], sum(report_dedications[project_name].values()))

        # let's write the projected and dedicated sum for that project
        _set_cell_value(report[f"F{PRJ_BRK_ROW}"], projects[project_name].projected)
        _set_cell_value(report[f"G{PRJ_BRK_ROW}"], 0)

        for month in range (1,13):
            time_to_dedicate = 0
            if month in report_dedications[project_name]:
                time_to_dedicate = report_dedications[project_name][month]
            _set_cell_value(report[f"{chr(ord('H') + month - 1)}{PRJ_BRK_ROW}"], time_to_dedicate)

    # Writing the TOTAL row
    total_row = PRJ_BRK_ROW+1+len(report_dedications.keys())
    cell = report[f"D{total_row}"]
    cell.value = f"=SUM($D{PRJ_BRK_ROW}:$D{PRJ_BRK_ROW+len(report_dedications.keys())-1})"
    cell.number_format = CELL_FMT
    for month in range (1,13):
        total_column = f"{chr(ord('H') + month - 1)}"
        cell = report[f"{total_column}{total_row}"]
        cell.value = f"=SUM(${total_column}{PRJ_BRK_ROW}:${total_column}{PRJ_BRK_ROW+len(report_dedications.keys())-1})"
        cell.number_format = CELL_FMT


    del workbook[defaults['worksheet-report-template']]
    workbook.save(defaults['project-report'])


def distribute_hours(args):
    MAX_DAILY_LIMIT = float(args.defaults['max-daily-limit'])*60
    MAX_YEARLY_LIMIT = float(args.defaults['max-yearly-limit'])*60
    granularity = float(args.justification_granularity)*60
    # Some sanity checks ...
    assert (MAX_DAILY_LIMIT >= granularity), f"Invalid configuration: daily limit ({int(MAX_DAILY_LIMIT)}') must be superior to justification granularity ({int(granularity)}')."
    assert (MAX_DAILY_LIMIT % granularity) == 0, f"Invalid configuration: daily limit ({int(MAX_DAILY_LIMIT)}') must be a multiple of the justification granularity ({int(granularity)}')."
    assert (MAX_YEARLY_LIMIT % granularity) == 0, f"Invalid configuration: yearly limit ({int(MAX_YEARLY_LIMIT)}') must be a multiple of the justification granularity ({int(granularity)}')."
    justification_start, justification_end = _get_justifiation_dates(args.justification_interval, args.justification_period)
    justification_year = justification_start.year 
    working_days, work_calendar = _get_working_days(justification_start, args.extra_holidays)

    # Load project breakdown for employee ...
    projects = _load_project_breakdown(args.project_breakdown, args.employee, args.only_projects, justification_year, working_days, args.defaults)
    
    logger.info(f"Justifying '{args.justification_interval}' interval from '{justification_start}' to '{justification_end}' ...")
    logger.debug(f"{str(justification_year)} has {len(working_days)} working day(s)")
    logger.debug(f"projects={pformat(list(projects.values()))}")

    # A map of working days with the project and the corresponding dedicated time
    daily_dedications = {}

    current_month = None
    for working_day in working_days:
        # I reset the project dedicated time when changing justification period / month
        if not current_month or current_month != working_day.month:
            logger.debug(f"New justification period detected. Resetting projects dedicated time ...")
            current_month = working_day.month
            list(map(lambda p: setattr(p, 'dedicated', 0), projects.values()))

        # List projects that still need to be justified for that day.
        # NOTE: that a project can be opened in the middle of the year.
        opened_projects = list(filter(lambda p: p.is_open(working_day) and not p.is_full(), projects.values()))
        if len(opened_projects) == 0:
            logger.debug(f"No project left opened on {working_day.strftime(DEFAULT_FMT)}")
            continue
        # daily_dedication corresponds to the time already allocated for that day
        daily_dedication = 0
        daily_limit = MAX_DAILY_LIMIT
        while daily_dedication < daily_limit:
            # Find the project that has the biggest time slot available for this day
            opened_project = sorted(opened_projects, key=lambda x: x.get_max_dedication(daily_limit - daily_dedication, granularity), reverse=True)[0]
            time_to_dedicate = opened_project.get_max_dedication(daily_limit - daily_dedication, granularity)
            if time_to_dedicate == 0:
                # XXX: Is that ever the case? Could a project be opened and have no time to dedicate?
                logger.debug(f"No time can be dedicated on {working_day.strftime(DEFAULT_FMT)}. Skipping to next day.")
                break
            opened_project.dedicated = opened_project.dedicated + time_to_dedicate
            logger.debug(f"Dedicating {time_to_dedicate}' to project {opened_project.name} on {working_day.strftime(DEFAULT_FMT)}")
            assert opened_project.dedicated <= opened_project.projected, f"Project dedication ({int(opened_project.dedicated)}') is superior to projected ({int(opened_project.projected)}') for project {opened_project.name}"
            daily_dedication += time_to_dedicate
            if daily_dedication > MAX_DAILY_LIMIT:
                break
            if working_day not in daily_dedications:
                daily_dedications[working_day] = {}
            daily_dedications[working_day][opened_project.name] = time_to_dedicate

    logger.debug(f"daily_dedications={pformat(daily_dedications)}")

    _generate_report(args.employee, args.overwrite, args.project_breakdown, projects, daily_dedications, justification_year, args.defaults)


def parse_command_line(defaults):
    parser = ArgumentParser(prog='distribute_hours',
                            description=__doc__, formatter_class=RawTextHelpFormatter)
    parser.add_argument(
        '--debug', action="store_true", help='Run the program in debug', required=False, default=False)
    parser.add_argument(
        '--quiet', action="store_true", help='Run the program in quiet mode', required=False, default=False)
    parser.add_argument(
        '--project-breakdown', help=f"The Excel file containing the project breakdown. By default '{defaults['project-breakdown']}'.", required=False, default=defaults['project-breakdown'])
    parser.add_argument(
        '--employee', help=f"The employee to justify hours for. By default '{defaults['employee']}'.", required=False, default=defaults['employee'])
    parser.add_argument(
        '--justification-period', type=lambda s: datetime.strptime(s, DEFAULT_FMT), help=f"The justification period in '{DEFAULT_FMT}' format. By default today's date.", required=False, default=date.today())
    parser.add_argument(
        '--justification-interval', help=f"The justification interval. Either 'monthly' or 'weekly'. By default '{defaults['justification-interval']}'.", required=False, default=defaults['justification-interval'])
    parser.add_argument(
        '--justification-granularity', type=float, help=f"The justification granularity in hour. By default '{defaults['justification-granularity']}'.", required=False, default=defaults['justification-granularity'])
    parser.add_argument(
        '--extra-holidays', help=f"A comma separated list of extra holidays to take into account (using '{DEFAULT_FMT}' format). First week of January by default.", required=False, default=defaults['extra-holidays'])
    parser.add_argument(
        '--only-projects', help=f"A comma separated list of projects to take into account. All project if not specified.", required=False, default=None)
    parser.add_argument(
        '--overwrite', action="store_true", help='Overwrite an existing report for the given employee', required=False, default=False)
    parser.set_defaults(func=distribute_hours)
    return parser.parse_args()


def main():
    with open(os.path.join(os.path.dirname(__file__), 'defaults.yaml'), 'r', encoding='utf-8') as f:
        defaults = yaml.safe_load(f)
    args = parse_command_line(defaults)
    args.defaults = defaults
    try:
        if args.debug:
            coloredlogs.install(level=logging.DEBUG, logger=logger)
        elif args.quiet:
            coloredlogs.install(level=logging.ERROR, logger=logger)
        else:
            coloredlogs.install(level=logging.INFO, logger=logger)
        return args.func(args)
    except Exception as e:
        logging.error(e, exc_info=True)
        return 1


if __name__ == '__main__':
    sys.exit(main())


