#!/usr/bin/env python

"""
pdn_project_allocation/solution.py

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

Solution class. Represents a potential solution. (The actual solving is done by
the Problem class.)

"""

import csv
import datetime
import logging
import operator
import os
from statistics import mean, median, variance
import sys
from typing import Dict, Generator, List, Tuple, TYPE_CHECKING

from cardinal_pythonlib.cmdline import cmdline_quote
from openpyxl.workbook.workbook import Workbook

from pdn_project_allocation.constants import (
    EXT_XLSX,
    SheetNames,
    VERSION,
    VERSION_DATE,
)
from pdn_project_allocation.helperfunc import (
    autosize_openpyxl_column,
    autosize_openpyxl_worksheet_columns,
)
from pdn_project_allocation.project import Project
from pdn_project_allocation.student import Student

if TYPE_CHECKING:
    from pdn_project_allocation.problem import Problem

log = logging.getLogger(__name__)


# =============================================================================
# Solution
# =============================================================================

class Solution(object):
    """
    Represents a potential solution.
    """
    def __init__(self,
                 problem: "Problem",
                 allocation: Dict[Student, Project]) -> None:
        """
        Args:
            problem:
                The :class:`Problem`, defining projects and students.
            allocation:
                The mapping of students to projects.
        """
        self.problem = problem
        self.allocation = allocation

    # -------------------------------------------------------------------------
    # Representations
    # -------------------------------------------------------------------------

    def __str__(self) -> str:
        """
        String representation.
        """
        lines = ["Solution:"]
        for student, project in self._gen_student_project_pairs():
            std = student.dissatisfaction(project)
            svd = project.dissatisfaction(student)
            lines.append(
                f"{student.shortname()} -> {project} "
                f"(student dissatisfaction {std}; "
                f"supervisor dissatisfaction {svd})")
        return "\n".join(lines)

    def shortdesc(self) -> str:
        """
        Very short description. Ordered by student number.
        """
        students = sorted(self.allocation.keys(), key=lambda s: s.number)
        parts = [f"{s.number}: {self.allocation[s].number}"
                 for s in students]
        return (
            "{" + ", ".join(parts) + "}" +
            f", student dissatisfaction {self.student_dissatisfaction_scores()}"
        )

    # -------------------------------------------------------------------------
    # Allocations
    # -------------------------------------------------------------------------

    def allocated_project(self, student: Student) -> Project:
        """
        Which project was allocated to this student?
        """
        return self.allocation[student]

    def allocated_students(self, project: Project) -> List[Student]:
        """
        Which students were allocated to this project?
        """
        return sorted(k for k, v in self.allocation.items() if v == project)

    def is_allocated(self, student: Student, project: Project) -> bool:
        """
        Is this student allocated to this project?
        """
        return self.allocation[student] == project

    def _gen_student_project_pairs(self) -> Generator[Tuple[Student, Project],
                                                      None, None]:
        """
        Generates ``student, project`` pairs in student order.
        """
        students = sorted(self.allocation.keys(), key=lambda s: s.number)
        for student in students:
            project = self.allocation[student]
            yield student, project

    # -------------------------------------------------------------------------
    # Student dissatisfaction
    # -------------------------------------------------------------------------

    def student_dissatisfaction_scores(self) -> List[float]:
        """
        All dissatisfaction scores.
        """
        dscores = []  # type: List[float]
        for student in self.problem.students:
            project = self.allocation[student]
            dscores.append(student.dissatisfaction(project))
        return dscores

    def student_dissatisfaction_median(self) -> float:
        """
        Median dissatisfaction per student.
        """
        return median(self.student_dissatisfaction_scores())

    def student_dissatisfaction_mean(self) -> float:
        """
        Mean dissatisfaction per student.
        """
        return mean(self.student_dissatisfaction_scores())

    def student_dissatisfaction_variance(self) -> float:
        """
        Variance of dissatisfaction scores.
        """
        return variance(self.student_dissatisfaction_scores())

    def student_dissatisfaction_min(self) -> float:
        """
        Minimum of dissatisfaction scores.
        """
        return min(self.student_dissatisfaction_scores())

    def student_dissatisfaction_max(self) -> float:
        """
        Maximum of dissatisfaction scores.
        """
        return max(self.student_dissatisfaction_scores())

    # -------------------------------------------------------------------------
    # Supervisor dissatisfaction
    # -------------------------------------------------------------------------

    def supervisor_dissatisfaction_scores_sum_students(self) -> List[float]:
        """
        All dissatisfaction scores. (If a project has several students, it
        scores the SUM of its dissatisfaction for each of those students
        scores.)
        """
        dscores = []  # type: List[float]
        for project in self.problem.projects:
            dscore = 0
            for student in self.problem.students:
                if self.allocation[student] == project:
                    dscore += project.dissatisfaction(student)
            dscores.append(dscore)
        return dscores

    def supervisor_dissatisfaction_scores_each_student(self) -> List[float]:
        """
        All dissatisfaction scores. (If a project has several students,
        multiple scores are returned for that project.)
        """
        dscores = []  # type: List[float]
        for project in self.problem.projects:
            for student in self.problem.students:
                if self.allocation[student] == project:
                    dscores.append(project.dissatisfaction(student))
        return dscores

    def supervisor_dissatisfaction_median(self) -> float:
        """
        Median dissatisfaction per student.
        """
        return median(self.supervisor_dissatisfaction_scores_each_student())

    def supervisor_dissatisfaction_mean(self) -> float:
        """
        Mean dissatisfaction per student.
        """
        return mean(self.supervisor_dissatisfaction_scores_each_student())

    def supervisor_dissatisfaction_variance(self) -> float:
        """
        Variance of dissatisfaction scores.
        """
        return variance(self.supervisor_dissatisfaction_scores_each_student())

    def supervisor_dissatisfaction_min(self) -> float:
        """
        Minimum of dissatisfaction scores.
        """
        return min(self.supervisor_dissatisfaction_scores_each_student())

    def supervisor_dissatisfaction_max(self) -> float:
        """
        Maximum of dissatisfaction scores.
        """
        return max(self.supervisor_dissatisfaction_scores_each_student())

    # -------------------------------------------------------------------------
    # Stability test
    # -------------------------------------------------------------------------

    def gen_better_projects(
            self,
            student: Student,
            project: Project) -> Generator[Project, None, None]:
        """
        Generates projects that this student prefers over the specified one.
        """
        for p in self.problem.gen_better_projects(student, project):
            yield p

    def gen_better_students(
            self,
            project: Project,
            student: Student) -> Generator[Student, None, None]:
        """
        Generates students that this project prefers over the specified one,
        for which they're eligible, AND who are are not already allocated to
        that project (bearing in mind that a project can have several
        students).
        """
        for s in self.problem.gen_better_students(project, student):
            if not self.is_allocated(s, project):
                yield s

    def stability(self, describe_all_failures: bool = True) -> Tuple[bool, str]:
        """
        Is the solution a stable match, and if not, why not? See README.rst for
        discussion. See also https://gist.github.com/joyrexus/9967709.

        Arguments:
            describe_all_failures:
                Show all reasons for failure.

        Returns:
            tuple: (stable, reason_for_instability)
        """
        stable = True
        instability_reasons = []  # type: List[str]
        for student, project in self._gen_student_project_pairs():
            for alt_project in self.gen_better_projects(student, project):
                for alt_proj_student in self.allocated_students(alt_project):
                    if alt_proj_student == student:
                        continue
                    if student in self.gen_better_students(alt_project,
                                                           alt_proj_student):
                        instability_reasons.append(
                            f"Pairing of student {student} to project "
                            f"{project} is unstable. "
                            f"The student would rather have alternative "
                            f"project {alt_project}, and that alternative "
                            f"project would rather have {student} than their "
                            f"current allocation of {alt_proj_student}."
                        )
                        stable = False
                        if not describe_all_failures:
                            return False, "\n\n".join(instability_reasons)
        if stable:
            return True, "[Stable]"
        else:
            return False, "\n\n".join(instability_reasons)

    def is_stable(self) -> bool:
        """
        Is the solution a stable match?
        """
        return self.stability(describe_all_failures=False)[0]

    # -------------------------------------------------------------------------
    # Saving
    # -------------------------------------------------------------------------

    def write_xlsx(self, filename: str) -> None:
        """
        Writes the solution to an Excel XLSX file (and its problem, for data
        safety).

        Args:
            filename:
                Name of file to write.
        """
        log.info(f"Writing output to: {filename}")

        # Not this way -- we can't then set column widths.
        #   wb = Workbook(write_only=True)  # doesn't create default sheet
        # Instead:
        wb = Workbook()
        wb.remove(wb.worksheets[0])

        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # Allocations, by student
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ss = wb.create_sheet(SheetNames.STUDENT_ALLOCATIONS)
        ss.append([
            "Student",
            "Project",
            "Supervisor",
            "Student's rank of allocated project (dissatisfaction score)",
        ])
        for student, project in self._gen_student_project_pairs():
            ss.append([
                student.name,
                project.title,
                project.supervisor_name(),
                student.dissatisfaction(project),
            ])
        autosize_openpyxl_worksheet_columns(ss)

        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # Allocations, by project
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ps = wb.create_sheet(SheetNames.PROJECT_ALLOCATIONS)
        ps.append([
            "Project",
            "Supervisor",
            "Student(s)",
            "Students' rank(s) of allocated project (dissatisfaction score)",
            "Project supervisor's rank(s) of allocated student(s) (dissatisfaction score)",  # noqa
        ])
        for project in self.problem.sorted_projects():
            student_names = []  # type: List[str]
            supervisor_dissatisfactions = []  # type: List[float]
            student_dissatisfactions = []  # type: List[float]
            for student in self.allocated_students(project):
                student_names.append(student.name)
                supervisor_dissatisfactions.append(
                    project.dissatisfaction(student)
                )
                student_dissatisfactions.append(
                    student.dissatisfaction(project)
                )
            ps.append([
                project.title,
                project.supervisor_name(),
                ", ".join(student_names),
                ", ".join(str(x) for x in student_dissatisfactions),
                ", ".join(str(x) for x in supervisor_dissatisfactions),
            ])
        autosize_openpyxl_worksheet_columns(ps)

        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # Popularity of projects
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        pp = wb.create_sheet(SheetNames.PROJECT_POPULARITY)
        pp.append([
            "Project",
            "Supervisor",
            "Total dissatisfaction score from all students",
            "Allocated student(s)",
            "Number of students expressing a preference",
            "Students expressing a preference",
        ])
        proj_to_unpop = {}  # type: Dict[Project, float]
        for project in self.problem.projects:
            unpopularity = 0
            for student in self.problem.students:
                unpopularity += student.dissatisfaction(project)
            proj_to_unpop[project] = unpopularity
        for project, unpopularity in sorted(proj_to_unpop.items(),
                                            key=operator.itemgetter(1, 0)):
            allocated_students = ", ".join(
                student.name
                for student in self.allocated_students(project)
            )
            student_prefs = {}  # type: Dict[Student, float]
            for student in self.problem.students:
                if student.preferences.actively_expressed_preference_for(
                        project):
                    student_prefs[student] = student.preferences.preference(
                        project)
            student_details = []  # type: List[str]
            for student, studpref in sorted(student_prefs.items(),
                                            key=operator.itemgetter(1, 0)):
                student_details.append(f"{student.name} ({studpref})")
            pp.append([
                project.title,
                project.supervisor_name(),
                unpopularity,
                allocated_students,
                len(student_details),
                ", ".join(student_details),
            ])

        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # Software, settings, and summary information
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        zs = wb.create_sheet(SheetNames.INFORMATION)
        is_stable, instability_reason = self.stability()
        zs_rows = [
            ["SOFTWARE DETAILS"],
            [],
            ["Software", "pdn_project_allocation"],
            ["Version", VERSION],
            ["Version date", VERSION_DATE],
            ["Source code",
             "https://github.com/RudolfCardinal/pdn_project_allocation"],
            ["Author", "Rudolf Cardinal (rudolf@pobox.com)"],
            [],
            ["RUN INFORMATION"],
            [],
            ["Date/time", datetime.datetime.now()],
            ["Overall weight given to student preferences",
             1 - self.problem.config.supervisor_weight],
            ["Overall weight given to supervisor preferences",
             self.problem.config.supervisor_weight],
            ["Command-line parameters", cmdline_quote(sys.argv)],
            ["Config", str(self.problem.config)],
            [],
            ["SUMMARY STATISTICS"],
            [],
            ["Student dissatisfaction median",
             self.student_dissatisfaction_median()],
            ["Student dissatisfaction mean",
             self.student_dissatisfaction_mean()],
            ["Student dissatisfaction variance",
             self.student_dissatisfaction_variance()],
            ["Student dissatisfaction minimum",
             self.student_dissatisfaction_min()],
            ["Student dissatisfaction minimum",
             self.student_dissatisfaction_max()],
            [],
            ["Supervisor dissatisfaction (with each student) median",
             self.supervisor_dissatisfaction_median()],
            ["Supervisor dissatisfaction (with each student) mean",
             self.supervisor_dissatisfaction_mean()],
            ["Supervisor dissatisfaction (with each student) variance",
             self.supervisor_dissatisfaction_variance()],
            ["Supervisor dissatisfaction (with each student) minimum",
             self.supervisor_dissatisfaction_min()],
            ["Supervisor dissatisfaction (with each student) maximum",
             self.supervisor_dissatisfaction_max()],
            [],
            ["Stable marriages?", str(is_stable)],
            ["If unstable, reason:", instability_reason]
        ]
        for row in zs_rows:
            zs.append(row)
        autosize_openpyxl_column(zs, 0)

        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # Problem definition
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        self.problem.write_to_xlsx_workbook(wb)

        wb.save(filename)
        wb.close()

    def write_data(self, filename: str) -> None:
        """
        Autodetects the file type from the extension and writes data to that
        file.
        """
        # File type?
        _, ext = os.path.splitext(filename)
        if ext == EXT_XLSX:
            self.write_xlsx(filename)
        else:
            raise ValueError(
                f"Don't know how to write file type {ext!r} for {filename!r}")

    def write_student_csv(self, filename: str) -> None:
        """
        Writes just the "per student" mapping to a CSV file, for comparisons
        (e.g. via ``meld``).
        """
        log.info(f"Writing student allocation data to: {filename}")
        with open(filename, "w") as file:
            writer = csv.writer(file)
            writer.writerow([
                "Student number",
                "Student name",
                "Project number",
                "Project name",
                "Student's rank of allocated project (dissatisfaction score)",
            ])
            for student, project in self._gen_student_project_pairs():
                writer.writerow([
                    student.number,
                    student.name,
                    project.number,
                    project.title,
                    student.dissatisfaction(project),
                ])
