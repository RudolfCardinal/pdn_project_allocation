#!/usr/bin/env python

"""
pdn_project_allocation/eligibility.py

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

Eligibility class.

"""

from collections import OrderedDict
import logging
from typing import List

from pdn_project_allocation.project import Project
from pdn_project_allocation.student import Student

log = logging.getLogger(__name__)


# =============================================================================
# Eligibility helpers
# =============================================================================

class Eligibility(object):
    """
    Simple wrapper around a map between students and projects.
    """

    def __init__(self,
                 students: List[Student],
                 projects: List[Project],
                 default_eligibility: bool = True,
                 allow_defunct_projects: bool = False) -> None:
        """
        Default constructor, which just sets default eligibility for everyone.

        Args:
            projects:
                All projects.
            students:
                All students.
            default_eligibility:
                Default value for "is student eligible for project"?
            allow_defunct_projects:
                Allow projects that permit no students?
        """
        self.students = sorted(students, key=lambda s: s.number)
        self.projects = sorted(projects, key=lambda p: p.number)
        self.eligibility = OrderedDict(
            (
                s,
                OrderedDict(
                    (p, default_eligibility)
                    for p in projects
                )
            )
            for s in students
        )
        self.allow_defunct_projects = allow_defunct_projects

    def __str__(self) -> str:
        """
        String representations.
        """
        if self.everyone_eligible_for_everything():
            return "All students eligible for all projects."
        lines = []  # type: List[str]
        for s, p_e in self.eligibility.items():
            projects_str = ", ".join(
                str(p)
                for p, e in p_e.items()
                if e
            )
            lines.append(f"{s}: eligible for {projects_str}")
        return "\n".join(lines)

    def assert_valid(self) -> None:
        """
        Perform internal checks, or raise an exception.
        """
        # 1. Every student has an eligible project.
        for s in self.students:
            assert any(self.is_eligible(s, p) for p in self.projects), (
                f"Error: student {s} is not eligible for any projects!"
            )
        # 2. Every project has an eligible student.
        for p in self.projects:
            if not any(self.is_eligible(s, p) for s in self.students):
                msg = f"Project {p} has no eligible students!"
                if self.allow_defunct_projects:
                    log.warning(msg)
                else:
                    raise AssertionError(
                        msg + " [If you meant this, set the "
                              "--allow_defunct_projects option.]")

    def set_eligibility(self,
                        student: Student,
                        project: Project,
                        eligible: bool):
        """
        Set eligibility for a specific student/project combination.

        Args:
            student: the student
            project: the project
            eligible: is the student eligible for the project?
        """
        self.eligibility[student][project] = eligible

    def is_eligible(self, student: Student, project: Project) -> bool:
        """
        Is the student eligible for the project?
        """
        return self.eligibility[student][project]

    def everyone_eligible_for_everything(self) -> bool:
        """
        Is this a simple problem in which everyone is eligible for everything?
        """
        return all(
            e
            for p_e in self.eligibility.values()
            for e in p_e.values()
        )
