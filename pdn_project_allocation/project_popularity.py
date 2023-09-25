#!/usr/bin/env python

"""
pdn_project_allocation/project_popularity.py

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
import operator
from typing import Any, Dict, List, Optional, TYPE_CHECKING

from scipy.stats import rankdata

from pdn_project_allocation.project import Project
from pdn_project_allocation.student import Student

if TYPE_CHECKING:
    from pdn_project_allocation.solution import Solution

log = logging.getLogger(__name__)


# =============================================================================
# ProjectPopularity
# =============================================================================


class ProjectPopularity:
    """
    Represents a project and its popularity information.
    """

    def __init__(self, project: Project, solution: "Solution") -> None:
        self.project = project
        self.solution = solution

        # Calculate unpopularity of this project
        problem = solution.problem
        self.unpopularity = 0.0
        for student in problem.students:
            self.unpopularity += student.dissatisfaction(project)

        self.satisfaction_rank = None  # type: Optional[float]

    @classmethod
    def headings(cls) -> List[str]:
        return [
            "Project",
            "Supervisor",
            "Total dissatisfaction score from all students",
            "Popularity rank",
            "Max. number of students",
            "Number of allocated student(s)",
            "Allocated student(s)",
            "Number of students expressing a preference",
            "Students expressing a preference († ineligible)",
        ]

    @classmethod
    def sort_and_assign_ranks(
        cls, popularities: List["ProjectPopularity"]
    ) -> None:
        """
        Modifies a list in place.
        """
        popularities.sort(key=lambda x: x.unpopularity)
        unpop = [x.unpopularity for x in popularities]
        ranks = rankdata(unpop, method="average")
        for i in range(len(popularities)):
            popularities[i].satisfaction_rank = ranks[i]

    def values(self) -> List[Any]:
        project = self.project
        eligibility = self.solution.problem.eligibility
        allocated_students = self.solution.allocated_students(project)
        student_prefs = {}  # type: Dict[Student, float]
        for student in self.solution.problem.students:
            if student.preferences.actively_expressed_preference_for(project):
                student_prefs[student] = student.preferences.preference(
                    project
                )
        student_details = []  # type: List[str]
        for student, studpref in sorted(
            student_prefs.items(), key=operator.itemgetter(1, 0)
        ):
            elig = "" if eligibility.is_eligible(student, project) else "†"
            student_details.append(f"{student.name} ({studpref}{elig})")

        return [
            self.project.title,
            self.project.supervisor_name(),
            self.unpopularity,
            self.satisfaction_rank,
            self.project.max_n_students,
            len(allocated_students),
            ", ".join(student.name for student in allocated_students),
            len(student_details),
            ", ".join(student_details),
        ]
