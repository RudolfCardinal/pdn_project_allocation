#!/usr/bin/env python

"""
pdn_project_allocation/constants.py

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

Constants and enums.

"""

from enum import Enum

from cardinal_pythonlib.enumlike import CaseInsensitiveEnumMeta


# =============================================================================
# Constants
# =============================================================================

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

DEFAULT_FONT = "Calibri"


class SheetNames:
    """
    Sheet names within the input/output spreadsheet file.
    """

    APPLIED_BUT_INELIGIBLE = "Applied_but_ineligible"
    ELIGIBILITY = "Eligibility"
    INFORMATION = "Information"  # output
    PROJECT_ALLOCATIONS = "Project_allocations"  # output
    PROJECT_POPULARITY = "Project_popularity"  # output
    PROJECTS = "Projects"  # input, output
    STUDENT_ALLOCATIONS = "Student_allocations"  # output
    STUDENT_PREFERENCES = "Student_preferences"  # input, output
    STUDENT_PREFERENCES_INTERNAL = "Student_preferences_internal"  # output
    SUPERVISOR_ALLOCATIONS = "Supervisor_allocations"  # output
    SUPERVISOR_PREFERENCES = "Supervisor_preferences"  # input, output
    SUPERVISOR_PREFERENCES_INTERNAL = "Supervisor_preferences_internal"  # out
    SUPERVISORS = "Supervisors"  # input, output
    UNALLOCATED_PROJECTS_WITH_CAPACITY = "Unallocated_projects_with_capacity"


class SheetHeadings:
    """
    Column headings within the input/output spreadsheets.
    """

    # Input:
    MAX_NUMBER_OF_PROJECTS = "Max_number_of_projects"
    MAX_NUMBER_OF_STUDENTS = "Max_number_of_students"
    PROJECT = "Project"
    SUPERVISOR = "Supervisor"

    # Additional for output:
    ELIGIBLE = "Eligible"
    N_PROJECTS_ALLOCATED = "N_projects_allocated"
    N_STUDENTS_ALLOCATED = "N_students_allocated"
    NOT_PREFERRED_PROJECT = "Project_not_preferred"
    NOT_PREFERRED_SUPERVISOR = "Supervisor_not_preferred"
    STUDENT = "Student"
    STUDENT_PREFERENCE = "Student_preference_rank"
    STUDENTS = "Student(s)"


class CsvHeadings:
    """
    Equivalently for simple CSV output.
    """

    DISSATISFACTION_SCORE = (
        "Students_rank_of_allocated_project_dissatisfaction_score"
    )
    PROJECT_NAME = "Project_name"
    PROJECT_NUMBER = "Project_number"
    STUDENT_NAME = "Student_name"
    STUDENT_NUMBER = "Student_number"


class SheetText:
    """
    Text used within some summary spreadsheets.
    """

    STUDENT_HAPPY = ""
    STUDENT_UNHAPPY_PROJECT = "Not a preferred project"
    STUDENT_UNHAPPY_SUPERVISOR = "Not a preferred supervisor"


class Switches:
    """
    Some switches are referred to in many places.
    """

    MISSING_ELIGIBILITY = "--missing_eligibility"
    STUDENT_MUST_HAVE_CHOICE = "--student_must_have_choice"


# =============================================================================
# Enum classes
# =============================================================================


class RankNotation(Enum, metaclass=CaseInsensitiveEnumMeta):
    """
    Ways of expressing ranks, and in particular ways of expressing tied ranks.
    See https://en.wikipedia.org/wiki/Ranking.
    """

    FRACTIONAL = "Fractional ranks (sum unaltered by ties; e.g. 1.5, 1.5, 3)"
    COMPETITION = "Standard competition ranks (e.g. 1, 1, 3)"
    DENSE = "Dense ranks (e.g. 1, 1, 2)"


class OptimizeMethod(Enum, metaclass=CaseInsensitiveEnumMeta):
    """
    Ways to solve our core problem mathematically.
    """

    MINIMIZE_DISSATISFACTION = (
        "Minimize weighted dissatisfaction "
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
        "Minimize weighted dissatisfaction, requiring stability if possible "
        "(as for MINIMIZE_DISSATISFACTION_STABLE), but falling back to "
        "unstable if not."
    )
    ABRAHAM_STUDENT = (
        "Abraham-Irving-Manlove 2007 (optimal for students); will not "
        "necessarily provide a solution"
    )
    ABRAHAM_SUPERVISOR = (
        "Abraham-Irving-Manlove 2007 (optimal for supervisors); will not "
        "necessarily provide a solution"
    )


DEFAULT_METHOD = OptimizeMethod.MINIMIZE_DISSATISFACTION_STABLE_FALLBACK
DEFAULT_RANK_NOTATION = RankNotation.FRACTIONAL
