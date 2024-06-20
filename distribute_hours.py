#!/usr/bin/env python
"""
Generate a justification sheet for the given employee and the given time period
"""

try:
    from dotenv import load_dotenv
    import collections
    import coloredlogs
    import json
    import logging
    import os
    import requests
    import sys
    import time
    import uuid
    import yaml
    from argparse import ArgumentParser, RawTextHelpFormatter
    from datetime import datetime
    from docx.shared import Mm
    from docxtpl import DocxTemplate, InlineImage, RichText
    from selenium import webdriver
    from lobe import LobeApplication
    from PIL import Image
    from workalendar.europe import Aragon
except ModuleNotFoundError as e:
    print(f"{e}. Did you load your environment?")
    sys.exit(1)

load_dotenv()
logger = logging.getLogger()

DEFAULT_FMT = '%Y-%m-%d'

def distribute_hours(args):
    # Read Project breakdown for this employee
    logger.info(f"Reading project breakdown for {args.employee} and ")
    project_breakdown = {}
    # Read Calendar
    calendar_aragon = Aragon()
    calendar_aragon.holidays(2024)
    # Create the yearly schedule on a 'justification-period' basis using the 'justification-granularity'
    for month in range (12):
        None

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
        '--justification-period', help=f"The Justification period. Either 'monthly' or 'weekly'. By default '{defaults['justification-period']}'.", required=False, default=defaults['justification-period'])
    parser.add_argument(
        '--justification-granularity', help=f"The granularity in minute of the justification. By default '{defaults['justification-granularity']}'.", required=False, default=defaults['justification-granularity'])
    parser.add_argument(
        '--output-dir', help=f"The directory where reports and screenshots should be stored.", required=False, default=defaults.get('output-dir',os.getcwd()))
    parser.set_defaults(func=distribute_hours)
    return parser.parse_args()


def main():
    with open(os.path.join(os.path.dirname(__file__), 'defaults.yaml'), 'r', encoding='utf-8') as f:
        defaults = yaml.safe_load(f)
    args = parse_command_line(defaults)
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


