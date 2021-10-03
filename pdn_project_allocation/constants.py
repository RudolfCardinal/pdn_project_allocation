#!/usr/bin/env python

"""
pdn_project_allocation/constants.py

===============================================================================

    Copyright (C) 2019-2021 Rudolf Cardinal (rudolf@pobox.com).

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

Constants and enums.

"""

from enum import Enum

from cardinal_pythonlib.enumlike import CaseInsensitiveEnumMeta


# =============================================================================
# Constants
# =============================================================================

VERSION = "1.4.0"
VERSION_DATE = "2021-10-03"

DEFAULT_PREFERENCE_POWER = 1.0
DEFAULT_MAX_SECONDS = 1e100  # the default in mip
DEFAULT_SUPERVISOR_WEIGHT = 0.3  # 70% student, 30% supervisor by default
RNG_SEED = 1234  # fixed
VERY_VERBOSE = False  # debugging option

EXT_XLSX = ".xlsx"
EXIT_FAILURE = 1
EXIT_SUCCESS = 0

INPUT_TYPES_SUPPORTED = [EXT_XLSX]
OUTPUT_TYPES_SUPPORTED = INPUT_TYPES_SUPPORTED

TRUE_VALUES = [1, "Y", "y", "T", "t"]
FALSE_VALUES = [0, "N", "n", "F", "f"]
MISSING_VALUES = ["", None]


# =============================================================================
# Enum classes
# =============================================================================

class SheetNames(object):
    """
    Sheet names within the input/output spreadsheet file.
    """
    ELIGIBILITY = "Eligibility"
    INFORMATION = "Information"  # output
    PROJECT_POPULARITY = "Project_popularity"  # output
    PROJECT_ALLOCATIONS = "Project_allocations"  # output
    PROJECTS = "Projects"  # input, output
    STUDENT_ALLOCATIONS = "Student_allocations"  # output
    STUDENT_PREFERENCES = "Student_preferences"  # input, output
    SUPERVISORS = "Supervisors"  # input, output
    SUPERVISOR_PREFERENCES = "Supervisor_preferences"  # input, output


class SheetHeadings(object):
    """
    Column headings within the input spreadsheet.
    """
    MAX_NUMBER_OF_PROJECTS = "Max_number_of_projects"
    MAX_NUMBER_OF_STUDENTS = "Max_number_of_students"
    PROJECT = "Project"
    SUPERVISOR = "Supervisor"


class OptimizeMethod(Enum, metaclass=CaseInsensitiveEnumMeta):
    MINIMIZE_DISSATISFACTION = (
        'Minimize weighted dissatisfaction '
        '(not necessarily requiring stable "marriages")'
    )
    MINIMIZE_DISSATISFACTION_STABLE_AB1996 = (
        "Minimize weighted dissatisfaction, requiring stability, "
        "via Abeledo & Blum (1996) method"
    )
    MINIMIZE_DISSATISFACTION_STABLE_CUSTOM = (
        "Minimize weighted dissatisfaction, requiring stability, "
        "via custom method that does not assume strict preferences"
    )
    MINIMIZE_DISSATISFACTION_STABLE = (
        "Minimize weighted dissatisfaction, requiring stability, "
        "via Abeledo & Blum (1996) falling back to custom method if required"
    )
    MINIMIZE_DISSATISFACTION_STABLE_FALLBACK = (
        "Minimize weighted dissatisfaction, requiring stability if possible"
        "(as for MINIMIZE_DISSATISFACTION_STABLE), but falling back to "
        "unstable if not."
    )
    ABRAHAM_STUDENT = "Abraham-Irving-Manlove 2007 (optimal for students)"
    ABRAHAM_SUPERVISOR = (
        "Abraham-Irving-Manlove 2007 (optimal for supervisors)"
    )


DEFAULT_METHOD = OptimizeMethod.MINIMIZE_DISSATISFACTION_STABLE_FALLBACK
