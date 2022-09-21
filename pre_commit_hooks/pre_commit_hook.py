#!/usr/bin/env python

r"""
pre_commit_hooks/pre_commit_hook.py

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

Pre-commit hook script that runs:

- black
- flake8

Usage:

.. code-block:: bash

    cd .git/hooks
    ln -s ../../pre_commit_hooks/pre_commit_hook.py pre-commit

This will run the ``pre_commit_hook.py`` script on each git commit (use ``git
commit -n`` to skip checks).

To avoid unexpected side effects, this script won't stash changes.
So if you have non-committed changes that break this you'll need to
stash your changes before committing.
"""

import logging
import os
from subprocess import CalledProcessError, PIPE, run
import sys
from typing import List

from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger


EXIT_FAILURE = 1

PRECOMMIT_DIR = os.path.dirname(os.path.realpath(__file__))
PROJECT_ROOT = os.path.join(PRECOMMIT_DIR, "..")
PYTHON_SOURCE_DIR = PROJECT_ROOT
FLAKE8_CONFIG_FILE = os.path.abspath(os.path.join(PROJECT_ROOT, "setup.cfg"))
GITHUB_ACTIONS_DIR = os.path.join(PROJECT_ROOT, ".github", "workflows")

log = logging.getLogger(__name__)


class PreCommitException(Exception):
    pass


def run_with_check(args: List[str]) -> None:
    run(args, check=True)


def check_python_style_and_errors() -> None:
    run_with_check(
        ["flake8", f"--config={FLAKE8_CONFIG_FILE}", PYTHON_SOURCE_DIR]
    )


def check_python_formatting() -> None:
    # Black does not support setup.cfg so we specify the options
    # on the command line (need to keep consistent with flake8)
    # TODO: Consider replacing setup.py and setup.cfg with pyproject.toml
    run_with_check(
        [
            "black",
            "--line-length",
            "79",
            "--diff",
            "--check",  # Should not need --exclude (reads .gitignore)
            PYTHON_SOURCE_DIR,
        ]
    )


# https://stackoverflow.com/questions/1871549/determine-if-python-is-running-inside-virtualenv
def in_virtualenv() -> bool:
    return hasattr(sys, "real_prefix") or (
        hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix
    )


def get_flake8_version() -> List[int]:
    command = ["flake8", "--version"]
    output = run(command, stdout=PIPE).stdout.decode("utf-8").split()[0]
    flake8_version = [int(n) for n in output.split(".")]

    return flake8_version


def main() -> None:
    if not in_virtualenv():
        log.error("pre_commit_hook.py must be run inside virtualenv")
        sys.exit(EXIT_FAILURE)

    log.info(f"Using flake8 config file {FLAKE8_CONFIG_FILE}")
    if not os.path.isfile(FLAKE8_CONFIG_FILE):
        log.error(f"Cannot find config file {FLAKE8_CONFIG_FILE}; aborting")
        sys.exit(EXIT_FAILURE)

    if get_flake8_version() < [3, 7, 8]:
        log.error(
            "flake8 version must be 3.7.8 or higher for type hint support"
        )
        sys.exit(EXIT_FAILURE)

    try:
        log.info("Checking Python formatting...")
        check_python_formatting()
        log.info("... done.")
        log.info("Checking for Python style and errors...")
        check_python_style_and_errors()
        log.info("... very stylish.")
    except CalledProcessError as e:
        log.error(str(e))
        log.error("Pre-commit hook failed. Check errors above")
        sys.exit(EXIT_FAILURE)


if __name__ == "__main__":
    main_only_quicksetup_rootlogger()
    main()
