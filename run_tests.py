#!/usr/bin/env python

"""
Run tests for pdn_project_allocation.
"""

import logging
import os
import subprocess

from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger

log = logging.getLogger(__name__)

THISDIR = os.path.dirname(os.path.realpath(__file__))
PROG = os.path.join(THISDIR, "pdn_project_allocation",
                    "pdn_project_allocation.py")
INPUTDIR = os.path.join(THISDIR, "testdata")
OUTPUTDIR = os.path.join(THISDIR, "testoutput")


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
    process("test1_equal_preferences_check_output_consistency.csv", "out1.csv")
    process("test2_trivial_perfect.csv", "out2.csv")
    process("test3_n60_two_equal_solutions.csv", "out3.csv")
    process("test4_n10_multiple_ties.csv", "out4.csv")
    process("test5_mean_vs_variance.csv", "out5.csv")
    process("test6_excel.xlsx", "out6.xlsx")


if __name__ == "__main__":
    main_only_quicksetup_rootlogger()
    main()
