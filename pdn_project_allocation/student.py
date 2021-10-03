#!/usr/bin/env python

"""
pdn_project_allocation/student.py

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

Student class.

"""

from typing import Dict, List, TYPE_CHECKING

from cardinal_pythonlib.reprfunc import auto_repr

from pdn_project_allocation.constants import DEFAULT_PREFERENCE_POWER
from pdn_project_allocation.preferences import Preferences

if TYPE_CHECKING:
    from pdn_project_allocation.project import Project


# =============================================================================
# Student
# =============================================================================

class Student(object):
    """
    Represents a single student, with their preferences.
    """
    def __init__(self,
                 name: str,
                 number: int,
                 preferences: Dict["Project", int],
                 n_projects: int,
                 allow_ties: bool = False,
                 preference_power: float = DEFAULT_PREFERENCE_POWER) -> None:
        """
        Args:
            name:
                Student's name.
            number:
                Row number of student (cosmetic only).
            preferences:
                Map from project to rank preference (1 to ``n_projects``
                inclusive).
            n_projects:
                Total number of projects (for validating inputs).
            allow_ties:
                Allow ties in preferences?
            preference_power:
                Power (exponent) to raise preferences to.
        """
        self.name = name
        self.number = number
        self.preferences = Preferences(
            n_options=n_projects,
            preferences=preferences,
            owner=self,
            allow_ties=allow_ties,
            preference_power=preference_power
        )

    def __str__(self) -> str:
        """
        String representation.
        """
        return f"{self.name} (St#{self.number})"

    def __repr__(self) -> str:
        return auto_repr(self)

    def description(self) -> str:
        """
        Verbose description.
        """
        return f"{self}: {self.preferences}"

    def shortname(self) -> str:
        """
        Name and number.
        """
        return f"{self.name} (St#{self.number})"

    def __lt__(self, other: "Student") -> bool:
        """
        Comparison for sorting, used for console display.
        Default sort is by case-insensitive name.
        """
        return self.name.lower() < other.name.lower()

    def dissatisfaction(self, project: "Project") -> float:
        """
        How dissatisfied is this student if allocated a particular project?
        """
        return self.preferences.preference(project)

    def exponentiated_dissatisfaction(self, project: "Project") -> float:
        """
        As for :meth:`dissatisfaction`, but raised to the desired power.
        """
        return self.preferences.exponentiated_preference(project)

    def explicitly_ranked_project(self, project: "Project") -> bool:
        """
        Did the student explicitly rank this project?
        """
        return self.preferences.actively_expressed_preference_for(project)

    def projects_in_descending_order(
            self, all_projects: List["Project"]) -> List["Project"]:
        """
        Returns projects in descending order of preference.
        """
        return self.preferences.items_descending_order(all_projects)
