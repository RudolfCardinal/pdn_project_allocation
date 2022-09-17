#!/usr/bin/env python

"""
pdn_project_allocation/project.py

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

Project class.

"""

import logging
from typing import Dict, List, Optional

from cardinal_pythonlib.reprfunc import auto_repr

from pdn_project_allocation.constants import DEFAULT_PREFERENCE_POWER
from pdn_project_allocation.helperfunc import supervisor_names_to_csv
from pdn_project_allocation.preferences import Preferences
from pdn_project_allocation.student import Student
from pdn_project_allocation.supervisor import Supervisor

log = logging.getLogger(__name__)


# =============================================================================
# Project
# =============================================================================


class Project(object):
    """
    Simple representation of a project.
    """

    def __init__(
        self,
        title: str,
        number: int,
        supervisors: List[Supervisor],
        max_n_students: int,
        allow_defunct_projects: bool = False,
    ) -> None:
        """
        Args:
            title:
                Project name.
            number:
                Project number (cosmetic only; matches input order).
            supervisors:
                The project's supervisor(s).
            max_n_students:
                Maximum number of students supported.
            allow_defunct_projects:
                Allow projects that permit no students?
        """
        assert title, "Missing project name"
        assert number >= 1, "Bad project number"
        if allow_defunct_projects:
            assert max_n_students >= 0, "Bad max_n_students"
        else:
            assert max_n_students >= 1, "Bad max_n_students"
        self.title = title
        self.number = number
        self.supervisors = supervisors
        self.max_n_students = max_n_students
        self.supervisor_preferences = None  # type: Optional[Preferences]
        # ... the project supervisor's preferences for students with respect
        #     to THIS project.

    def __str__(self) -> str:
        """
        String representation.
        """
        return f"{self.title} (P#{self.number})"

    def __repr__(self) -> str:
        return auto_repr(self)

    def __lt__(self, other: "Project") -> bool:
        """
        Comparison for sorting, used for console display.
        Default sort is by case-insensitive name.
        """
        return self.title.lower() < other.title.lower()

    def description(self) -> str:
        """
        Describes the project.
        """
        return (
            f"{self} (max {self.max_n_students} students): "
            f"{self.supervisor_preferences}"
        )

    def set_supervisor_preferences(
        self,
        n_students: int,
        preferences: Dict[Student, int],
        allow_ties: bool = False,
        preference_power: float = DEFAULT_PREFERENCE_POWER,
    ) -> None:
        """
        Sets the supervisor's preferences about students for a project.
        """
        self.supervisor_preferences = Preferences(
            n_options=n_students,
            owner=self,
            preferences=preferences,
            allow_ties=allow_ties,
            preference_power=preference_power,
        )

    def dissatisfaction(self, student: Student) -> float:
        """
        How dissatisfied is this project's supervisor if allocated a particular
        student?
        """
        return self.supervisor_preferences.preference(student)

    def exponentiated_dissatisfaction(self, project: "Project") -> float:
        """
        As for :meth:`dissatisfaction`, but raised to the desired power.
        """
        return self.supervisor_preferences.exponentiated_preference(project)

    def students_in_descending_order(
        self, all_students: List[Student]
    ) -> List[Student]:
        """
        Returns students in descending order of preference.
        """
        return self.supervisor_preferences.items_descending_order(all_students)

    def supervisor_name(self) -> str:
        """
        Name of this project's supervisor(s), in human-readable CSV format.
        """
        names = [s.name for s in self.supervisors]
        return supervisor_names_to_csv(names)

    def is_supervised_by(self, supervisor: Supervisor) -> bool:
        """
        Is this project supervised by this particular supervisor?
        """
        return supervisor in self.supervisors

    def at_least_one_supervisor_has_a_project_cap(self) -> bool:
        """
        Does at least one supervisor of this project? have a cap on the number
        of projects they can take?
        """
        return any(s.max_n_projects is not None for s in self.supervisors)
