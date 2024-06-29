#!/usr/bin/env python
"""
Generate a justification sheet for an employee and a time period.

The process is composed of different steps:
    1. It starts with a project breakdown that lists, for each projet, the 
       employee projected dedication by year.
    2. This yearly projected dedication is then refined into a projected 
       dedication split into the given justification interval. For instance if
       the justification interval is set to 'month' the yearly projected 
       dedication will be split evenly into 12 intervals. Conversly if the 
       justification interval is set to 'week' the yearly projected 
       dedication will be split evenly into 52 intervals. Note that the 
       refined projected dedication can be adjusted by the user and picked 
       up by the script using the '--use-refined-projected-dedication' 
       parameter.
    3. Once the projected dedication has been refined for each justification
       interval, the script will start filling in each day of said interval 
       with the projects' projected dedication that are opened for that
       specific day, taking care of not breaching the limits imposed at 
       different level (yearly, monthly, by project, etc.).
    4. Once the dedication has been established for the justification interval
       the script generate a report with the details of the justification for 
       the given justification interval, as well as a daily justification 
       sheet

NOTE: The script relies heavily on the fact that the input excel file has a 
      valid format. Do not modify the input Excel format without changing the
      corresponding parameter in the 'defauls.yaml' file.

NOTE: Use the 'resources/project-breakdown.xlsx' as a reference for the input 
      Excel file. 

SYNOPSIS:

# Generate a pristine justification sheet for John Smith for the current month
?> distribute_hours.py --employee "John Smith"

# Generate a pristine justification sheet for John Smith for June 2023
?> distribute_hours.py --employee "John Smith" --justification-period 2023-06-24

"""

try:
    from datetime import date, datetime, timedelta
    from dotenv import load_dotenv
    from enum import Enum
    import time
    import calendar
    import coloredlogs
    import logging
    import math
    import os
    import re
    import sys
    import yaml
    from argparse import ArgumentParser, RawTextHelpFormatter
    from pprint import pprint, pformat
    from workalendar.europe import Aragon
    import openpyxl
    from openpyxl.formatting import Rule
    from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
    from openpyxl.styles import NamedStyle, Font, Border, Side, Color, PatternFill, Protection
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.cell.cell import Cell
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.utils import quote_sheetname, absolute_coordinate
except ModuleNotFoundError as e:
    print(f"{e}. Did you load your environment?")
    sys.exit(1)

load_dotenv()
logger = logging.getLogger()

DEFAULT_FMT = '%Y-%m-%d'
CELL_FMT = "0.0" # "[h]:mm"
PREFERRED_FMT = 'decimal'
MONTH_INTERVAL = 'month'
WEEK_INTERVAL = 'week'


# The first row containing project information
PRJ_1ST_ROW = 7
PRJ_NAME_COL = 'A'
# The column containing the projected dedication for the given project
PROJECTED_COL = 'B'
# The column containing the Excel SUM of justified dedication for the given project
JUSTIFIED_COL = 'C'
# The column containing the justified dedication for the given month
MONTH_1ST_COL = 'D'
MONTH_1ST_COL_NUM = ord(MONTH_1ST_COL) - ord('A') + 1
MONTH_LST_COL = chr(ord(MONTH_1ST_COL) + 12 - 1)
MONTH_LST_COL_NUM = ord(MONTH_LST_COL) - ord('A') + 1


class Project:
    """Base class to all projects"""

    def __init__(self, name:str, start:date, end:date):
        self.name = name
        self.name_coord = None
        self.start = start
        self.end = end
        self.interval_limits = []
        self.projected = 0
        self.projected_coord = None
        self.justified = 0

    def __str__(self) -> str:
        return self.__repr__()
    
    def __repr__(self) -> str:
        return f"<PRJ {self.name} ({self.start.strftime(DEFAULT_FMT)} -> {self.end.strftime(DEFAULT_FMT)}, projected={self.projected}, justified={self.justified})>"
    
    def is_open(self, working_day:date) -> bool:
        """Returns whether a working day falls within a project

        Args:
            working_day (date): the day to test

        Returns:
            bool: True if the working day falls within a project
        """
        return self.start <= working_day and working_day <= self.end
    
    def is_full(self, current_month:int) -> bool:
        """Returns whether the project has used up all its projected dedication

        Returns:
            bool: True if no more time can be dedicated to the project
        """
        assert 1 <= current_month and current_month <= 12, "current_month must be a datetime month (from 1 to 12)"
        return ((self.justified >= self.projected) or (self.justified >= self.interval_limits[current_month-1]))
    
    def get_max_dedication(self, available_today:int, granularity:int, current_month:int) -> int:
        """Returns the maximum number of minute that can be dedicated by the project taking into account different limits.

        Args:
            limit_daily (int): the maximum number of minutes that can be daily dedicated to all projects
            granularity (int): the minimum interval for justification
            current_month (int): the 

        Case: return project.available 
        +--------------> interval_limits
        +------------>   available_today
        +         +      granular_span
        +--->            project.available

        Case: return project.available 
        +--------------> interval_limits
        +----->          available_today
        +         +      granular_span
        +--->            project.available

        Case: return available_today
        +--------------> interval_limits
        +----->          available_today
        +         +      granular_span
        +----------->    project.available

        Case: return granular_span
        +--------------> interval_limits
        +------------>   available_today
        +         +      granular_span
        +----------->    project.available

        Case: return granular_span
        +--------------> interval_limits
        +------------>   available_today
        +         +      granular_span
        +----------->    project.available

        Case: return interval_limits
        +------->        interval_limits
        +------------>   available_today
        +         +      granular_span
        +----------->    project.available
        
        Returns:
            int: the number of minutes
        """
        assert 1 <= current_month and current_month <= 12, "current_month must be a datetime month (from 1 to 12)"
        available_this_year = self.projected - self.justified
        available_this_month = self.interval_limits[current_month-1] - self.justified
        granular_span = max(math.ceil(min([available_this_year, available_this_month, available_today]) / granularity), 1) * granularity
        return min([available_this_year, available_this_month, available_today, granular_span])

def _get_cell_coordinate(worksheet, cell_range:str) -> str:
    return f"{quote_sheetname(worksheet.title)}!{absolute_coordinate(cell_range)}"


def _load_justification_hints(worksheet:Worksheet) -> dict:
    hints = {}
    # Look up the project name in column PRJ_NAME_COL from PRJ_1ST_ROW
    for col in worksheet.iter_cols(min_col=1, max_col=1, min_row=PRJ_1ST_ROW):
        for prj_cell in col:
            project_name = prj_cell.value
            if not project_name:
                break
            hints[project_name] = {}
            for month_col in worksheet.iter_cols(min_col=MONTH_1ST_COL_NUM, 
                                                 max_col=MONTH_LST_COL_NUM, 
                                                 min_row=prj_cell.row, 
                                                 max_row=prj_cell.row):
                for month_cell in month_col:
                    # NOTE: we work with minutes...
                    hints[project_name][month_cell.column-MONTH_1ST_COL_NUM] = round(month_cell.value * 60)
    logger.debug(f"hints={pformat(hints)}")
    return hints


def _load_project_breakdown(project_breakdown_filename:str, 
                            employee:str, 
                            only_projects:str, 
                            justification_interval:str, 
                            use_justification_hints:bool, 
                            justification_year:int, 
                            working_days:list, 
                            defaults) -> dict:
    """Returns a dict containing information about the projects for the given employee

    Args:
        project_breakdown_filename (str): the filename of the Excel file containing the project breakdown
        employee (str): the name of the employee
        only_projects (str): a comma separated list of project name to filter on.
        justification_interval (str): indicates whether justification interval is 'month' or 'week'
        use_justification_hints (bool): whether to use the hints found in a possible already existing justification sheet
        justification_year (int): the year for which the justification is to be produced.
        working_days (list(date)): a sorted list of working days as per the established working calendar.
        defaults (dict): the default read from the 'default.yaml' file

    Raises:
        ValueError: raised if the employee does not have a project breakdown

    Returns:
        A dict containg the projects details:
        { project_name -> Project (project_name, start, end, interval_limits, justified, projected)

    """
    logger.info(f"Loading projects breakdown from {project_breakdown_filename} for '{employee}' ...")
    workbook = openpyxl.load_workbook(project_breakdown_filename)
    if employee not in workbook:
        raise ValueError(f"Invalid '--employee' parameter: No worksheet '{employee}' found in {project_breakdown_filename}.")
    
    hints = {}
    if use_justification_hints:
        hint_worksheet_name = f"{employee} - {justification_year}" 
        if hint_worksheet_name in workbook:
            logger.debug(f"Will use justification hints found in worksheet '{hint_worksheet_name}' ...")
            hints = _load_justification_hints(workbook[hint_worksheet_name])
        else:
            logger.warning(f"You requested to use hints, but no worksheet named '{hint_worksheet_name}' was found. Skipping hints altogether.")

    worksheet = workbook[employee]
    # determining the column indexes for each relevant column
    name_col_idx = 0
    start_col_idx = 0
    end_col_idx = 0
    year_columns = {}
    # I determine the column indices...
    logger.debug(f"Determining column indices ...")
    for col in worksheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            if cell.value == defaults['project-name-column']:
                name_col_idx = cell.column
            elif cell.value == defaults['project-start-column']:
                start_col_idx = cell.column
            elif cell.value == defaults['project-end-column']:
                end_col_idx = cell.column
            elif (isinstance(cell.value, int) or isinstance(cell.value, str)) and int(cell.value) >= 2000 and int(cell.value) <= 2100:
                # The column seems to be a year...
                year_columns[int(cell.value)] = cell.column
            else:
                logger.warning(f"Ignoring unknown header {cell.column} (='{cell.value}' /{type(cell.value)})")
    # The map { project_name -> Project }
    logger.debug(f"Extracting projects details ...")
    projects = { }
    only_projects = only_projects.split(",") if only_projects else None
    # Each row is a project ...
    for i in range(worksheet.min_row+1, worksheet.max_row+1):
        name_cell = worksheet.cell(row=i, column=name_col_idx)
        
        if only_projects and name_cell.value not in only_projects:
            logger.warning(f"Skipping project '{name_cell.value}' ...")
            continue

        logger.debug(f"Loading project '{name_cell.value}' ...")
        start_cell = worksheet.cell(row=i, column=start_col_idx)
        end_cell = worksheet.cell(row=i, column=end_col_idx)
        projected_cell = worksheet.cell(row=i, column=year_columns[justification_year])
        if not projected_cell.value:
            projected_cell.value = 0.0
        assert type(start_cell.value) == type(datetime.now()), f"Invalid type detected for cell '{employee}!{start_cell.coordinate}'. Expected a date."
        assert type(end_cell.value) == type(datetime.now()), f"Invalid type detected for cell '{employee}!{end_cell.coordinate}'. Expected a date."
        assert type(projected_cell.value) == type(0.0) or type(projected_cell.value) == type(0), f"Invalid type detected for cell '{employee}!{projected_cell.coordinate}'. Expected a float or an int."

        # Let's create the project
        project = Project(str(name_cell.value), start_cell.value, end_cell.value)
        project.name_coord = _get_cell_coordinate(worksheet, name_cell.coordinate)
        projected = projected_cell.value
        project.projected_coord = _get_cell_coordinate(worksheet, projected_cell.coordinate)
        # The project has some projected dedication for the justification_year
        if projected > 0:
            project.projected = round(projected * 60)
            project_working_days = list(filter(lambda d: project.start <= d and d <= project.end, working_days))
            if len(project_working_days) <= 0:
                raise ValueError (f"Invalid project breakdown: project '{project.name}' has projected dedication ({int(project.projected)}'), but no working days. It is configured to start on the {project.start.strftime(DEFAULT_FMT)} and finish on the {project.end.strftime(DEFAULT_FMT)}.")
            if justification_interval == MONTH_INTERVAL:
                number_of_months = project_working_days[-1].month - project_working_days[0].month + 1
                project.interval_limits = [0] * 12
                default_interval_limit = int(float(project.projected) / number_of_months)
                for j, _ in enumerate(project.interval_limits):
                    working_day = project.start.replace(year=justification_year, month=j+1)
                    if not project.is_open(working_day):
                        logger.debug(f"Project '{project.name}' is not opened on {working_day.strftime(DEFAULT_FMT)}. Skipping hints ...")
                        continue
                    hint = int(hints.get(project.name, {}).get(j, default_interval_limit))
                    if hint != default_interval_limit:
                        logger.debug(f"Overriding default dedication for project {project.name} for {working_day.strftime(DEFAULT_FMT)}: from {default_interval_limit} to {hint}")
                    project.interval_limits[j] = hint
            elif justification_interval == WEEK_INTERVAL:
                assert False, "NOT IMPLEMENTED"
            else:
                raise ValueError(f"Invalid '--justification-interval' parameter. Must be either 'month' or 'week'.")
        projects[project.name] = project

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
    if interval == MONTH_INTERVAL:
        start = period.replace(day=1)
        end = period.replace(day=calendar.monthrange(start.year, start.month)[1])
    else:
        start = period - timedelta(days=period.weekday())
        end = period + timedelta(days=6-period.weekday()) 
    return start, end


def _get_working_days(justification_start, extra_holidays):
    """Returns the list of working day for the justification year removing the specified extra holidays

    Args:
        justification_start (datetime.date): the start of the justification period. Usually the 1st of the month to justify
        extra_holidays (list(datetime.date)): a list of date containing the employee holidays

    Returns:
        list(date): the list of working days
    """
    # ok, I always justify a full year and then extract the intersting period.
    # This is due to the fact that we do not want to justify periods that are not a multiple of the 
    # justification granulariry. Only during the last period of the year, is this allowed.
    logger.debug(f"Creating the holiday calendar for Aragon in {justification_start.year} ...")
    work_calendar = Aragon()
    # With the holiday calendar I create the list of the opened days in the justification year
    extra_holidays = extra_holidays.split(",")
    january_first = justification_start.replace(month=1, day=1)

    working_days = []
    for delta_day in (range(1, 367 if calendar.isleap(justification_start.year) else 366)):
        target_day = january_first + timedelta(days=delta_day)
        if work_calendar.is_working_day(target_day) and target_day.strftime(DEFAULT_FMT) not in extra_holidays:
            working_days.append(target_day)
    return working_days, work_calendar


def _set_cell_to_time(cell:Cell, minutes:float):
    """Set the value of the given cell 

    Args:
        cell (Cell): _description_
        minutes (float): _description_
    """
    if PREFERRED_FMT == 'decimal':
        cell.value = minutes/60 
        cell.number_format = "0.0"
    else:
        # Excel processes time entries as a decimal fraction of a day
        cell.value = minutes/(60 * 24)
        cell.number_format = "[h]:mm"


def _set_cell_to_project_name(cell:Cell, project:str, defaults:dict):
    if defaults['project-report-use-reference']:
        cell.value = f"={project.name_coord}"
    else:
        cell.value = project.name
    cell.protection = Protection(locked=True)
    try:
        cell.style = defaults['project-report-project-style']
    except:
        logger.warning(f"Invalid configuration 'project-report-project-style': Missing named style '{defaults['project-report-project-style']}' in Excel file. Check your defaults.yaml.")


def _set_cell_to_projected_dedication(cell:Cell, project:Project, defaults:dict):
    if defaults['project-report-use-reference']:
        cell.value = f"={project.projected_coord}"
    else:
        _set_cell_to_time(cell, project.projected)
    cell.number_format = CELL_FMT
    cell.protection = Protection(locked=True)


def _set_cell_to_justified_dedication(cell:Cell, named_range_name:str, defaults:dict):
    cell.value = f"=SUM({named_range_name})"
    cell.number_format = CELL_FMT


def _set_cell_to_total(cell:Cell, named_range_name:str, defaults:dict):
    cell.value = f"=SUM({named_range_name})"
    cell.number_format = CELL_FMT
    

def _safe_named_range_name(name:str) -> str:
    return re.sub('[^0-9a-zA-Z_]+', '', name).lower()


def _add_named_range(worksheet, name:str, cell_range:str) -> str:
    named_range = DefinedName(_safe_named_range_name(name), attr_text=_get_cell_coordinate(worksheet, cell_range))
    worksheet.defined_names.add(named_range)
    return named_range


def _get_month_col(month_1st_col:str, month:int) -> int:
    return chr(ord(month_1st_col) + month - 1)


def _generate_report(employee, project_breakdown:str, projects:dict, working_days, daily_dedications, justification_year:int, defaults:dict):
    logger.info("Generating report ...")
    workbook = openpyxl.load_workbook(project_breakdown, keep_vba=True)
    if defaults['worksheet-report-template'] not in workbook:
        raise ValueError(f"Invalid configuration: No worksheet '{defaults['worksheet-report-template']}' found in {project_breakdown}.")
    report_template = workbook[defaults['worksheet-report-template']]

    report_title = f"{employee} - {justification_year}"
    if report_title in workbook:
        logger.warning(f"Removing existing report '{report_title}' ...")
        del workbook[report_title]

    report = workbook.copy_worksheet(report_template)
    report.title = report_title

    # Should try to use defined name instead...
    # definition = report.defined_names["employee"]
    report['B1'].value = employee
    report['B2'].value = employee
    report['B3'].value = justification_year
    report['B4'].value = defaults['max-yearly-limit']

    # Let's define some ranges
    _add_named_range(report, "employee", 'B2')
    _add_named_range(report, "justification_year", 'B3')
    _add_named_range(report, "max_yearly_limit", 'B4')

    # I massage the daily_dedications as { project_name -> {month -> justified} }
    justifications = {}
    for working_day, project_dedications in daily_dedications.items():
        for project_name, justified in project_dedications.items():
            if project_name not in justifications.keys():
                justifications[project_name] = {}
            if working_day.month not in justifications[project_name]:
                justifications[project_name][working_day.month] = 0
            justifications[project_name][working_day.month] += justified

    logger.debug(f"justifications={pformat(justifications)}")
    # TODO: Way too much magic numbers! Indices are way too fragile! How to improve templating?


    for i, project_name in enumerate(sorted(justifications.keys())):

        project_row = PRJ_1ST_ROW + i
        project = projects[project_name]

        report.insert_rows(project_row)

        # let's set a couple of ranges
        _add_named_range(report, f"prj_total_projected_{project_name}", f"{PROJECTED_COL}{project_row}")
        _add_named_range(report, f"prj_total_justified_{project_name}", f"{JUSTIFIED_COL}{project_row}")
        _add_named_range(report, f"prj_monthly_justified_{project_name}", f"{MONTH_1ST_COL}{project_row}:{MONTH_LST_COL}{project_row}")

        # Set the project name
        _set_cell_to_project_name(report[f"{PRJ_NAME_COL}{project_row}"], project, defaults)

        # The 'Horas propuestas' for that project
        _set_cell_to_projected_dedication(report[f"{PROJECTED_COL}{project_row}"], project, defaults)

        # The 'Horas justificadas' for that project (utilizamos la SUM de Excel como comprobacion)
        _set_cell_to_justified_dedication(report[f"{JUSTIFIED_COL}{project_row}"], _safe_named_range_name(f"prj_monthly_justified_{project_name}"), defaults)

        # The justified time for each month
        for month in range (1,13):
            justified = 0
            if month in justifications[project_name]:
                justified = justifications[project_name][month]
            month_column = _get_month_col(MONTH_1ST_COL, month)
            _set_cell_to_time(report[f"{month_column}{project_row}"], justified)

    #
    # 'Total proyectos' row
    #
    # 'Total proyectos' / 'Horas propuestas'
    totals_row = project_row+2
    _add_named_range(report, "all_prj_projected", f"{PROJECTED_COL}{PRJ_1ST_ROW}:{PROJECTED_COL}{project_row}")
    _add_named_range(report, "all_prj_projected_total", f"{PROJECTED_COL}{totals_row}")
    _set_cell_to_total(report[f"{PROJECTED_COL}{totals_row}"], "all_prj_projected", defaults)

    # 'Total proyectos' / 'Horas justificadas'
    _add_named_range(report, "all_prj_justified", f"{JUSTIFIED_COL}{PRJ_1ST_ROW}:{JUSTIFIED_COL}{project_row}")
    _add_named_range(report, "all_prj_justified_total", f"{JUSTIFIED_COL}{totals_row}")
    _set_cell_to_total(report[f"{JUSTIFIED_COL}{totals_row}"], "all_prj_justified", defaults)

    # 'Total proyectos' / Month
    for month in range (1,13):
        month_column = _get_month_col(MONTH_1ST_COL, month)
        named_range_name = _safe_named_range_name(f"all_prj_monthly_justified_{month}")
        _add_named_range(report, named_range_name, f"{month_column}{PRJ_1ST_ROW}:{month_column}{project_row}")
        _set_cell_to_total(report[f"{month_column}{totals_row}"], named_range_name, defaults)

    #
    # 'Limites mensuales' row
    #
    # We write the maximum number of hours for this month MAX_DAILY_LIMIT*monthly_working_days
    MAX_DAILY_LIMIT = float(defaults['max-daily-limit'])*60
    limits_row = totals_row+1
    for month in range (1,13):
        month_column = _get_month_col(MONTH_1ST_COL, month)
        monthly_working_days = list(filter(lambda d: d.month == month, working_days))
        monthly_limit = MAX_DAILY_LIMIT * len(monthly_working_days)
        _set_cell_to_time(report[f"{month_column}{limits_row}"], monthly_limit)

    alert_color=defaults['project-report-alert-color']
    warning_color=defaults['project-report-warning-color']
    ok_color=defaults['project-report-ok-color']
    grey_color=defaults['project-report-grey-color']

    # Doing some more formatting
    # Set the project's monthly dedication in grey if it is 0
    fill = PatternFill(bgColor=grey_color)
    rule = CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{MONTH_1ST_COL}{PRJ_1ST_ROW}:{MONTH_LST_COL}{project_row}", rule)

    # BEWARE: The conditional formatting syntax is slightly different from the formula syntax. It uses ',' instead of ';'
    # Set the monthly justified dedication in red/yellow/green if it breaches the monthly limit (row offset = 1)
    fill = PatternFill(bgColor=alert_color)
    rule = FormulaRule(formula=['INDIRECT("RC",FALSE) > OFFSET(INDIRECT("RC",FALSE),1,0)'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{MONTH_1ST_COL}{totals_row}:{MONTH_LST_COL}{totals_row}", rule)

    fill = PatternFill(bgColor=warning_color)
    rule = FormulaRule(formula=['INDIRECT("RC",FALSE) = OFFSET(INDIRECT("RC",FALSE),1,0)'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{MONTH_1ST_COL}{totals_row}:{MONTH_LST_COL}{totals_row}", rule)

    fill = PatternFill(bgColor=ok_color)
    rule = FormulaRule(formula=['INDIRECT("RC",FALSE) < OFFSET(INDIRECT("RC",FALSE),1,0)'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{MONTH_1ST_COL}{totals_row}:{MONTH_LST_COL}{totals_row}", rule)

    # Set the project's total justified dedication in red/yellow/green if it breaches the projected limit (col offset = -1)
    fill = PatternFill(bgColor=alert_color)
    rule = FormulaRule(formula=['INDIRECT("RC",FALSE) > OFFSET(INDIRECT("RC",FALSE),0,-1)'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{JUSTIFIED_COL}{PRJ_1ST_ROW}:{JUSTIFIED_COL}{project_row}", rule)

    fill = PatternFill(bgColor=warning_color)
    rule = FormulaRule(formula=['INDIRECT("RC",FALSE) = OFFSET(INDIRECT("RC",FALSE),0,-1)'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{JUSTIFIED_COL}{PRJ_1ST_ROW}:{JUSTIFIED_COL}{project_row}", rule)

    fill = PatternFill(bgColor=ok_color)
    rule = FormulaRule(formula=['INDIRECT("RC",FALSE) < OFFSET(INDIRECT("RC",FALSE),0,-1)'], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{JUSTIFIED_COL}{PRJ_1ST_ROW}:{JUSTIFIED_COL}{project_row}", rule)

    # Set the total projected dedication in red if it breaches the yearly limit, green otherwise
    fill = PatternFill(bgColor=alert_color)
    rule = FormulaRule(formula=[f"${PROJECTED_COL}${totals_row} > max_yearly_limit"], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{PROJECTED_COL}{totals_row}", rule)
    fill = PatternFill(bgColor=ok_color)
    rule = FormulaRule(formula=[f"${PROJECTED_COL}${totals_row} <= max_yearly_limit"], stopIfTrue=True, fill=fill)
    report.conditional_formatting.add(f"{PROJECTED_COL}{totals_row}", rule)

    del workbook[defaults['worksheet-report-template']]
    workbook.save(defaults['project-report'])


def distribute_hours(args):
    # We interanlly work with minutes.
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
    projects = _load_project_breakdown(args.project_breakdown, 
        args.employee, 
        args.only_projects, 
        args.justification_interval, 
        args.use_justification_hints,
        justification_year, 
        working_days,
        args.defaults)
    
    logger.info(f"Justifying '{args.justification_interval}' interval from '{justification_start}' to '{justification_end}' ...")
    logger.debug(f"{str(justification_year)} has {len(working_days)} working day(s)")
    logger.debug(f"projects={pformat(list(projects.values()))}")

    # A map of working days with the project and the corresponding justified time
    daily_dedications = {}

    current_month = None
    for working_day in working_days:
        # I reset the project justified dedication time when changing justification period / month
        if not current_month or current_month != working_day.month:
            logger.debug(f"New justification period detected. Resetting projects justified dedication ...")
            current_month = working_day.month
            list(map(lambda p: setattr(p, 'justified', 0), projects.values()))

        # List projects that still need to be justified for that day.
        # NOTE: that a project can be opened in the middle of the year.
        if working_day.month == 10:
            logger.debug(f"projects={pformat(list(projects.values()))}")
        opened_projects = list(filter(lambda p: p.is_open(working_day) and not p.is_full(current_month), projects.values()))
        if len(opened_projects) == 0:
            logger.debug(f"No project left opened on {working_day.strftime(DEFAULT_FMT)}")
            continue
        # daily_dedication corresponds to the time already allocated for that day
        daily_dedication = 0
        daily_limit = MAX_DAILY_LIMIT
        while daily_dedication < daily_limit:
            # Find the project that has the biggest time slot available for this day
            opened_project = sorted(opened_projects, key=lambda x: x.get_max_dedication(daily_limit - daily_dedication, granularity, current_month), reverse=True)[0]
            time_to_dedicate = opened_project.get_max_dedication(daily_limit - daily_dedication, granularity, current_month)
            if time_to_dedicate == 0:
                # XXX: Is that ever the case? Could a project be opened and have no time to dedicate?
                logger.debug(f"No time can be dedicated on {working_day.strftime(DEFAULT_FMT)}. Skipping to next day.")
                break
            opened_project.justified = opened_project.justified + time_to_dedicate
            logger.debug(f"Dedicating {time_to_dedicate}' to project {opened_project.name} on {working_day.strftime(DEFAULT_FMT)}")
            assert opened_project.justified <= opened_project.projected, f"Project justified dedication ({int(opened_project.justified)}') is superior to projected dedication ({int(opened_project.projected)}') for project {opened_project.name}"
            daily_dedication += time_to_dedicate
            if daily_dedication > MAX_DAILY_LIMIT:
                break
            if working_day not in daily_dedications:
                daily_dedications[working_day] = {}
            daily_dedications[working_day][opened_project.name] = time_to_dedicate

    logger.debug(f"daily_dedications={pformat(daily_dedications)}")

    _generate_report(args.employee, args.project_breakdown, projects, working_days, daily_dedications, justification_year, args.defaults)


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
        '--employee', help=f"The employee to justify hours for. Note that the project breakdown file must contain a worksheet whose name is equal to the employee's name. By default '{defaults['employee']}'.", required=False, default=defaults['employee'])
    parser.add_argument(
        '--justification-period', type=lambda s: datetime.strptime(s, DEFAULT_FMT), help=f"The justification period in '{DEFAULT_FMT}' format. By default today's date.", required=False, default=date.today().strftime(DEFAULT_FMT))
    parser.add_argument(
        '--justification-interval', help=f"The justification interval. Either 'month' or 'week'. By default '{defaults['justification-interval']}'.", required=False, default=defaults['justification-interval'])
    parser.add_argument(
        '--justification-granularity', type=float, help=f"The justification granularity in hour. By default '{defaults['justification-granularity']}'.", required=False, default=defaults['justification-granularity'])
    parser.add_argument(
        '--extra-holidays', help=f"A comma separated list of extra holidays to take into account (using '{DEFAULT_FMT}' format). First week of January by default.", required=False, default=defaults['extra-holidays'])
    parser.add_argument(
        '--only-projects', help=f"A comma separated list of projects to take into account. All project if not specified.", required=False, default=None)
    parser.add_argument(
        '--use-justification-hints', action="store_true", help='Use a previously created justification worksheet to calculate projected dedication. The justified dedication found in said worksheet will take precedence on the projected dedication. The hint worksheet must have been created by this program in order to be readable.', required=False, default=False)
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


