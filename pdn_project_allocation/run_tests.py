#!/usr/bin/env python

"""
pdn_project_allocation/run_tests.py

Run tests for pdn_project_allocation.
"""

import logging
import os
import subprocess
from typing import List

from cardinal_pythonlib.cmdline import cmdline_quote
from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger

log = logging.getLogger(__name__)

THISDIR = os.path.dirname(os.path.realpath(__file__))
PROG = os.path.join(THISDIR, "pdn_project_allocation.py")
INPUTDIR = os.path.join(THISDIR, "testdata")
OUTPUTDIR = os.path.join(THISDIR, os.pardir, "testoutput")


def process(infile: str, outfile: str,
            other_options: List[str] = None) -> None:
    cmdargs = [
        "python",
        PROG,
        os.path.join(INPUTDIR, infile),
        "--output", os.path.join(OUTPUTDIR, outfile),
        "--verbose",

        # "--method", "minimize_dissatisfaction_stable",
        # "--method", "minimize_dissatisfaction_stable_custom",
        "--method", "minimize_dissatisfaction_stable_fallback",
        # "--method", "abraham_student",

        # "--power", "3.0",
    ]
    if other_options:
        cmdargs += other_options
    log.warning(cmdline_quote(cmdargs))
    subprocess.check_call(cmdargs)


def main() -> None:
    main_only_quicksetup_rootlogger()
    process("test1_equal_preferences_check_output_consistency.xlsx",
            "test_out1.xlsx")
    process("test2_trivial_perfect.xlsx", "test_out2.xlsx")
    process("test3_n60_two_equal_solutions.xlsx", "test_out3.xlsx")
    process("test4_n10_multiple_ties.xlsx", "test_out4.xlsx")
    process("test5_mean_vs_variance.xlsx", "test_out5.xlsx")
    process("test6_multiple_students_per_project.xlsx", "test_out6.xlsx")
    process("test7_eligibility.xlsx", "test_out7.xlsx")
    process("test8_tied_preferences.xlsx", "test_out8.xlsx",
            ["--allow_supervisor_preference_ties"])
    process("test9_supervisor_constraints.xlsx", "test_out9.xlsx",
            ["--debug_model", "--no_shuffle"])


if __name__ == "__main__":
    main()
