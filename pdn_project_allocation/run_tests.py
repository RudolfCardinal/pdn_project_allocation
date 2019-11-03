#!/usr/bin/env python

"""
pdn_project_allocation/run_tests.py

Run tests for pdn_project_allocation.
"""

import logging
import os
import subprocess

from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger

log = logging.getLogger(__name__)

THISDIR = os.path.dirname(os.path.realpath(__file__))
PROG = os.path.join(THISDIR, "pdn_project_allocation.py")
INPUTDIR = os.path.join(THISDIR, "testdata")
OUTPUTDIR = os.getcwd()


def process(infile: str, outfile: str) -> None:
    cmdargs = [
        "python",
        PROG,
        os.path.join(INPUTDIR, infile),
        "--output", os.path.join(OUTPUTDIR, outfile),
        # "--power", "3.0",
    ]
    log.info(repr(cmdargs))
    subprocess.check_call(cmdargs)


def main() -> None:
    process("test1_equal_preferences_check_output_consistency.xlsx", "out1.xlsx")
    process("test2_trivial_perfect.xlsx", "out2.xlsx")
    process("test3_n60_two_equal_solutions.xlsx", "out3.xlsx")
    process("test4_n10_multiple_ties.xlsx", "out4.xlsx")
    process("test5_mean_vs_variance.xlsx", "out5.xlsx")
    process("test6_excel.xlsx", "out6.xlsx")


if __name__ == "__main__":
    main_only_quicksetup_rootlogger()
    main()
