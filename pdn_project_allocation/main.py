#!/usr/bin/env python

"""
pdn_project_allocation/main.py

===============================================================================

    Copyright (C) 2019 Rudolf Cardinal (rudolf@pobox.com).

    This file is part of pdn_project_allocation.

    This is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This software is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this software. If not, see <https://www.gnu.org/licenses/>.

===============================================================================

Command-line entry point.

"""

import argparse
import distutils.util
import logging
import random
import sys
import traceback

from cardinal_pythonlib.argparse_func import (
    RawDescriptionArgumentDefaultsHelpFormatter,
)
from cardinal_pythonlib.enumlike import keys_descriptions_from_enum
from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger
from cardinal_pythonlib.cmdline import cmdline_quote

from pdn_project_allocation.config import Config
from pdn_project_allocation.constants import (
    DEFAULT_MAX_SECONDS,
    DEFAULT_METHOD,
    DEFAULT_PREFERENCE_POWER,
    DEFAULT_SUPERVISOR_WEIGHT,
    EXIT_FAILURE,
    EXIT_SUCCESS,
    FALSE_VALUES,
    INPUT_TYPES_SUPPORTED,
    OptimizeMethod,
    OUTPUT_TYPES_SUPPORTED,
    RNG_SEED,
    SheetHeadings,
    SheetNames,
    TRUE_VALUES,
)
from pdn_project_allocation.problem import Problem

log = logging.getLogger(__name__)


# =============================================================================
# main
# =============================================================================


def main() -> None:
    """
    Command-line entry point.
    """
    # noinspection PyTypeChecker
    parser = argparse.ArgumentParser(
        formatter_class=RawDescriptionArgumentDefaultsHelpFormatter,
        description=f"""
Allocate students to projects, maximizing some version of happiness.

The input spreadsheet should have the following format (in each case, the
first row is the title row):

    Sheet name:
        {SheetNames.SUPERVISORS}
    Description:
        List of supervisors (one per row) and their project/student capacity.
        Capacity values can be left blank for "no maximum".
    Format:
        {SheetHeadings.SUPERVISOR}      {SheetHeadings.MAX_NUMBER_OF_PROJECTS}  {SheetHeadings.MAX_NUMBER_OF_STUDENTS}
        Dr Smith        3                       5
        Dr Jones
        Dr Lucas                                2
        ...             ...                     ...
        
    Sheet name:
        {SheetNames.PROJECTS}
    Description:
        List of projects (one per row), their student capacity, and their
        supervisor. If a project has multiple supervisors, separate them with
        commas in the {SheetHeadings.SUPERVISOR!r} column (order is not
        important).
    Format:
        {SheetHeadings.PROJECT}         {SheetHeadings.MAX_NUMBER_OF_STUDENTS}  {SheetHeadings.SUPERVISOR}
        Project One     1                       Dr Jones
        Project Two     1                       Dr Jones
        Project Three   2                       Dr Smith
        Project Four    2                       Dr Smith, Dr Jones
        ...             ...                     ...
        
    Sheet name:
        {SheetNames.STUDENT_PREFERENCES}
    Description:
        List of students (one per row) and their rank preferences (1 = top, 2 =
        next, etc.) for projects (one per column).
    Format:
        <ignored>       Project One     Project Two     Project Three   ...
        Miss Smith      1               2                               ...
        Mr Jones        2               1               3               ...
        ...             ...             ...             ...             ...
    
    Sheet name:
        {SheetNames.ELIGIBILITY}
    Description:
        OPTIONAL sheet, showing which students are eligible for which
        projects. If absent, all students are eligible for all projects.
        Use {TRUE_VALUES} for "eligible".
        Use {FALSE_VALUES} for "ineligible".
        Use --missing_eligibility to control the handling of empty cells.
    Format:
        <ignored>       Project One     Project Two     Project Three   ...
        Miss Smith      1               1               1               ...
        Mr Jones        1               0               1               ...
        ...             ...             ...             ...             ...
    
    Sheet name:
        {SheetNames.SUPERVISOR_PREFERENCES}
    Description:
        List of projects (one per column) and their supervisor's rank
        preferences (1 = top, 2 = next, etc.) for students (one per row). If
        a project has multiple supervisors, those supervisors must agree their
        joint "project preference" (these preferences are project-specific,
        not supervisor-specific).
    Format:
        <ignored>       Project One     Project Two     Project Three   ...
        Miss Smith      1               1                               ...
        Mr Jones        2               2                               ...
        ...             ...             ...

""",  # noqa
    )
    parser.add_argument("--verbose", action="store_true", help="Be verbose")

    file_group = parser.add_argument_group("Files")
    file_group.add_argument(
        "filename",
        type=str,
        help="Spreadsheet filename to read. "
        "Input file types supported: " + str(INPUT_TYPES_SUPPORTED),
    )
    file_group.add_argument(
        "--output",
        type=str,
        help="Optional filename to write output to. "
        "Output types supported: " + str(OUTPUT_TYPES_SUPPORTED),
    )
    file_group.add_argument(
        "--output_student_csv",
        type=str,
        help="Optional filename to write student CSV output to.",
    )

    data_group = parser.add_argument_group("Data")
    data_group.add_argument(
        "--allow_student_preference_ties",
        action="store_true",
        help="Allow students to express tied preferences "
        "(e.g. 2.5 for joint second/third place)?",
    )
    data_group.add_argument(
        "--allow_supervisor_preference_ties",
        action="store_true",
        help="Allow supervisors to express tied preferences "
        "(e.g. 2.5 for joint second/third place)?",
    )
    data_group.add_argument(
        "--missing_eligibility",
        dest="missing_eligibility",
        type=lambda x: bool(distutils.util.strtobool(x)),
        default=None,
        # https://stackoverflow.com/questions/15008758/parsing-boolean-values-with-argparse  # noqa
        help="If an eligibility cell is blank, treat it as eligible (use "
        "'True'/'yes'/1 etc.) or ineligible (use 'False'/'no'/0 etc.)? "
        "Default, of None, means empty cells are invalid.",
    )
    data_group.add_argument(
        "--allow_defunct_projects",
        action="store_true",
        help="Allow projects that say that all students are ineligible (e.g. "
        "because they've been pre-allocated by different process)?",
    )

    method_group = parser.add_argument_group("Method")
    method_group.add_argument(
        "--supervisor_weight",
        type=float,
        default=DEFAULT_SUPERVISOR_WEIGHT,
        help="Weight allocated to supervisor preferences (student preferences "
        "are weighted as [1 minus this])",
    )
    method_group.add_argument(
        "--preference_power",
        type=float,
        default=DEFAULT_PREFERENCE_POWER,
        help="Power (exponent) to raise preferences by.",
    )
    method_group.add_argument(
        "--student_must_have_choice",
        action="store_true",
        help="Prevent students being allocated to projects they've not "
        "explicitly ranked?",
    )

    technical_group = parser.add_argument_group("Technicalities")
    technical_group.add_argument(
        "--maxtime",
        type=float,
        default=DEFAULT_MAX_SECONDS,
        help="Maximum time (in seconds) to run MIP optimizer for",
    )
    technical_group.add_argument(
        "--seed",
        type=int,
        default=None,
        help="Seed for random number generator. "
        "DO NOT USE FOR ACTUAL ALLOCATIONS; IT IS UNFAIR (because it "
        "tempts the operator to re-run with different seeds). "
        "FOR DEBUGGING USE ONLY.",
    )
    technical_group.add_argument(
        "--no_shuffle",
        action="store_true",
        help="Don't shuffle anything. FOR DEBUGGING USE ONLY.",
    )
    technical_group.add_argument(
        "--debug_model",
        action="store_true",
        help="Report the details of the MIP model before solving.",
    )
    method_k, method_desc = keys_descriptions_from_enum(
        OptimizeMethod, keys_to_lower=True
    )
    method_group.add_argument(
        "--method",
        type=str,
        choices=method_k,
        default=DEFAULT_METHOD.name,
        help=f"Method of solving. -- {method_desc} --",
    )

    args = parser.parse_args()
    main_only_quicksetup_rootlogger(
        level=logging.DEBUG if args.verbose else logging.INFO
    )

    # Seed RNG
    if args.seed is not None:
        log.warning(
            "You have specified --seed. FOR DEBUGGING USE ONLY: "
            "THIS IS NOT FAIR FOR REAL ALLOCATIONS!"
        )
        seed = args.seed
    else:
        seed = RNG_SEED
    random.seed(seed)

    # Go
    config = Config(
        allow_defunct_projects=args.allow_defunct_projects,
        allow_student_preference_ties=args.allow_student_preference_ties,
        allow_supervisor_preference_ties=args.allow_supervisor_preference_ties,
        cmd_args=vars(args),
        debug_model=args.debug_model,
        filename=args.filename,
        max_time_s=args.maxtime,
        missing_eligibility=args.missing_eligibility,
        no_shuffle=args.no_shuffle,
        optimize_method=OptimizeMethod[args.method],
        preference_power=args.preference_power,
        student_must_have_choice=args.student_must_have_choice,
        supervisor_weight=args.supervisor_weight,
    )
    log.info(f"Command: {cmdline_quote(sys.argv)}")
    log.info(f"Config: {config}")
    problem = Problem.read_data(config)
    if args.output:
        log.debug(problem)
    else:
        log.info(problem)
    solution = problem.best_solution()
    if solution:
        if args.output:
            log.debug(solution)
        else:
            log.info(solution)
        if args.output:
            solution.write_data(args.output)
        else:
            log.warning(
                "Output not saved. Specify the --output option for that."
            )
        if args.output_student_csv:
            solution.write_student_csv(args.output_student_csv)
        sys.exit(EXIT_SUCCESS)
    else:
        log.error("No solution found!")
        sys.exit(EXIT_FAILURE)


if __name__ == "__main__":
    try:
        main()
    except Exception as _top_level_exception:
        log.critical(str(_top_level_exception))
        log.critical(traceback.format_exc())
        sys.exit(EXIT_FAILURE)
