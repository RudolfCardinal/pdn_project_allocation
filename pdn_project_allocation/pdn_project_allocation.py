#!/usr/bin/env python

"""
pdn_project_allocation/pdn_project_allocation.py

See README.rst

"""

import argparse
from collections import OrderedDict
import datetime
import logging
import os
import random
from statistics import mean, variance
from typing import (Any, Dict, Generator, List, Optional, Sequence,
                    Tuple, Union)

from cardinal_pythonlib.argparse_func import RawDescriptionArgumentDefaultsHelpFormatter  # noqa
from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger
from cardinal_pythonlib.maths_py import sum_of_integers_in_inclusive_range
from openpyxl.cell import Cell
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from mip import BINARY, minimize, Model, xsum

log = logging.getLogger(__name__)

VERSION = "1.0.1"
VERSION_DATE = "2020-09-11"

ALMOST_ONE = 0.99
DEFAULT_MAX_SECONDS = 60
DEFAULT_SUPERVISOR_WEIGHT = 0.3  # 70% student, 30% supervisor by default
RNG_SEED = 1234  # fixed
VERY_VERBOSE = False  # debugging option

EXT_XLSX = ".xlsx"

INPUT_TYPES_SUPPORTED = [EXT_XLSX]
OUTPUT_TYPES_SUPPORTED = INPUT_TYPES_SUPPORTED


# =============================================================================
# Playing with the mip package
# =============================================================================

r"""

Just for fun, the n-queens problem from
https://python-mip.readthedocs.io/en/latest/examples.html:

from sys import stdout
from mip import Model, xsum, MAXIMIZE, BINARY

# number of queens
n = 75

queens = Model()

x = [[queens.add_var('x({},{})'.format(i, j), var_type=BINARY)
      for j in range(n)] for i in range(n)]

# one per row
for i in range(n):
    queens += xsum(x[i][j] for j in range(n)) == 1, 'row({})'.format(i)

# one per column
for j in range(n):
    queens += xsum(x[i][j] for i in range(n)) == 1, 'col({})'.format(j)

# diagonal \
for p, k in enumerate(range(2 - n, n - 2 + 1)):
    queens += xsum(x[i][j] for i in range(n) for j in range(n)
                   if i - j == k) <= 1, 'diag1({})'.format(p)

# diagonal /
for p, k in enumerate(range(3, n + n)):
    queens += xsum(x[i][j] for i in range(n) for j in range(n)
                   if i + j == k) <= 1, 'diag2({})'.format(p)

queens.optimize()

text = ""
if queens.num_solutions:
    for i, v in enumerate(queens.vars):
        text += 'Q ' if v.x >= 0.99 else '. '
        if i % n == n-1:
            text += "\n"

print(text)

"""


# =============================================================================
# Enum classes
# =============================================================================

class SheetNames(object):
    """
    Sheet names within the input/output spreadsheet file.
    """
    PROJECTS = "Projects"  # input, output
    STUDENT_PREFERENCES = "Student_preferences"  # input, output
    SUPERVISOR_PREFERENCES = "Supervisor_preferences"  # input, output
    STUDENT_ALLOCATIONS = "Student_allocations"  # output
    PROJECT_ALLOCATIONS = "Project_allocations"  # output
    INFORMATION = "Information"  # output


class SheetHeadings(object):
    """
    Column headings within the input spreadsheet.
    """
    PROJECT_NAME = "Project_name"
    MAX_NUMBER_OF_STUDENTS = "Max_number_of_students"


# =============================================================================
# Preferences
# =============================================================================

class Preferences(object):
    """
    Represents preference as a mapping from arbitrary objects (being preferred)
    to ranks.
    """
    def __init__(self,
                 n_options: int,
                 preferences: Dict[Any, int] = None,
                 owner: Any = None) -> None:
        """
        Args:
            n_options:
                Total number of things to be judged.
            preferences:
                Mapping from "thing being judged" to "rank preference" (1 is
                best).
            owner:
                Person/thing expressing preferences (for cosmetic purposes
                only).

        Other attributes:
        - ``available_dissatisfaction``: sum of [1 ... ``n_options`]
        - ``allocated_dissatisfaction``: sum of expressed preference ranks.
          (For example, if you only pick your top option, with rank 1, then you
          have expressed a total dissatisfaction of 1. If you have expressed
          a preference for rank #1 and rank #2, you have expressed a total
          dissatisfaction of 3.)
        """
        self._n_options = n_options
        self._preferences = OrderedDict()  # type: Dict[Any, int]
        self._owner = owner
        self._available_dissatisfaction = sum_of_integers_in_inclusive_range(
            1, n_options)
        self._allocated_dissatisfaction = 0

        if preferences:
            for item, rank in preferences.items():
                if rank is not None:
                    self.add(item, rank, _validate=False)
                    # ... defer validation until all data in...
            self._validate()  # OK, now validate

    def __str__(self) -> str:
        """
        String representation.
        """
        parts = ", ".join(f"{k} → {v}" for k, v in self._preferences.items())
        return (
            f"Preferences({parts}; "
            f"unranked options score {self.unranked_item_dissatisfaction})"
        )

    def set_n_options(self, n_options: int) -> None:
        """
        Sets the total number of options, and ensures that the preferences
        are compatible with this.
        """
        self._n_options = n_options
        self._validate()

    def add(self, item: Any, rank: int, _validate: bool = True) -> None:
        """
        Add a preference for an item.

        Args:
            item:
                Thing for which a preference is being assessed.
            rank:
                Integer preference rank (1 best, 2 next, etc.).
            _validate:
                Validate immediately?
        """
        assert item not in self._preferences, (
            f"Can't add same item twice; attempt to re-add {item!r}"
        )
        assert isinstance(rank, int), (
            f"Only integer preferences allowed at present; was {rank!r}"
        )
        assert rank not in self._preferences.values(), (
            f"No duplicate dissatisfaction scores allowed at present: "
            f"attempt to re-add rank {rank}"
        )
        self._preferences[item] = rank
        self._allocated_dissatisfaction += rank
        if _validate:
            self._validate()

    def _validate(self) -> None:
        """
        Validates:

        - that there are some options;
        - that preferences for all options are in the range [1, ``n_options``];
        - that the ``allocated_dissatisfaction`` is no more than the
          ``available_dissatisfaction``.

        Raises:
            :exc:`AssertionError` upon failure.
        """
        assert self._n_options > 0, "No options"
        for rank in self._preferences.values():
            assert 1 <= rank <= self._n_options, f"Invalid preference: {rank!r}"
        assert (
            self._allocated_dissatisfaction <= self._available_dissatisfaction
        ), (
            "Dissatisfaction scores add up to more than the maximum"
        )

    @property
    def unallocated_dissatisfaction(self) -> int:
        """
        The amount of available "dissatisfaction", not yet allocated to an
        item (see :class:`Preferences`).
        """
        return self._available_dissatisfaction - self._allocated_dissatisfaction

    @property
    def unranked_item_dissatisfaction(self) -> Optional[float]:
        """
        The mean "dissatisfaction" (see :class:`Preferences`) for every option
        without an explicit preference, or ``None`` if there are no such
        options.
        """
        n_unranked = self._n_options - len(self._preferences)
        return (
            self.unallocated_dissatisfaction / n_unranked
            if n_unranked > 0 else None
        )

    def preference(self, item: Any) -> Union[int, float]:
        """
        Returns a numerical preference score for an item. Will return the
        "unranked" item dissatisfaction if no preference has been expressed for
        this particular item.

        Args:
            item:
                The item to look up.
        """
        return self._preferences.get(item, self.unranked_item_dissatisfaction)

    def raw_preference(self, item: Any) -> Optional[int]:
        """
        Returns the raw preference for an item (for reproducing the input).

        Args:
            item:
                The item to look up.
        """
        return self._preferences.get(item)  # returns None if absent


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
                 n_projects: int) -> None:
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
        """
        self.name = name
        self.number = number
        self.preferences = Preferences(
            n_options=n_projects,
            preferences=preferences,
            owner=self,
        )

    def __str__(self) -> str:
        """
        String representation.
        """
        return f"{self.name} (S#{self.number})"

    def description(self) -> str:
        """
        Verbose description.
        """
        return f"{self}: {self.preferences}"

    def shortname(self) -> str:
        """
        Name and number.
        """
        return f"{self.name} (#{self.number})"

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


# =============================================================================
# Project
# =============================================================================

class Project(object):
    """
    Simple representation of a project.
    """
    def __init__(self,
                 name: str,
                 number: int,
                 max_n_students: int) -> None:
        """
        Args:
            name:
                Project name.
            number:
                Project number (cosmetic only; matches input order).
            max_n_students:
                Maximum number of students supported.
        """
        assert name, "Missing name"
        assert number >= 1, "Bad project number"
        assert max_n_students >= 1, "Bad max_n_students"
        self.name = name
        self.number = number
        self.max_n_students = max_n_students
        self.supervisor_preferences = None  # type: Optional[Preferences]

    def __str__(self) -> str:
        """
        String representation.
        """
        return f"{self.name} (P#{self.number})"

    def __lt__(self, other: "Project") -> bool:
        """
        Comparison for sorting, used for console display.
        Default sort is by case-insensitive name.
        """
        return self.name.lower() < other.name.lower()

    def description(self) -> str:
        """
        Describes the project.
        """
        return (
            f"{self} (max {self.max_n_students} students): "
            f"{self.supervisor_preferences}"
        )

    def set_supervisor_preferences(self,
                                   n_students: int,
                                   preferences: Dict[Student, int]) -> None:
        """
        Sets the supervisor's preferences about students for a project.
        """
        self.supervisor_preferences = Preferences(
            n_options=n_students,
            owner=self,
            preferences=preferences
        )

    def dissatisfaction(self, student: Student) -> float:
        """
        How dissatisfied is this project's supervisor if allocated a particular
        student?
        """
        return self.supervisor_preferences.preference(student)


# =============================================================================
# Solution
# =============================================================================

class Solution(object):
    """
    Represents a potential solution.
    """
    def __init__(self,
                 problem: "Problem",
                 allocation: Dict[Student, Project],
                 supervisor_weight: float) -> None:
        """
        Args:
            problem:
                The :class:`Problem`, defining projects and students.
            allocation:
                The mapping of students to projects.
        """
        self.problem = problem
        self.allocation = allocation
        self.supervisor_weight = supervisor_weight

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
        lines.append("")
        lines.append(f"Student dissatisfaction mean: "
                     f"{self.student_dissatisfaction_mean()}")
        lines.append(f"Student dissatisfaction variance: "
                     f"{self.student_dissatisfaction_variance()}")
        lines.append(f"Supervisor dissatisfaction mean: "
                     f"{self.supervisor_dissatisfaction_mean()}")
        lines.append(f"Supervisor dissatisfaction variance: "
                     f"{self.supervisor_dissatisfaction_variance()}")
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

    def _gen_student_project_pairs(self) -> Generator[Tuple[Student, Project],
                                                      None, None]:
        """
        Generates ``student, project`` pairs in student order.
        """
        students = sorted(self.allocation.keys(), key=lambda s: s.number)
        for student in students:
            project = self.allocation[student]
            yield student, project

    def student_dissatisfaction_scores(self) -> List[float]:
        """
        All dissatisfaction scores.
        """
        dscores = []  # type: List[float]
        for student in self.problem.students:
            project = self.allocation[student]
            dscores.append(student.dissatisfaction(project))
        return dscores

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

    def supervisor_dissatisfaction_scores(self) -> List[float]:
        """
        All dissatisfaction scores.
        """
        dscores = []  # type: List[float]
        for project in self.problem.projects:
            dscore = 0
            for student in self.problem.students:
                if self.allocation[student] == project:
                    dscore += project.dissatisfaction(student)
            dscores.append(dscore)
        return dscores

    def supervisor_dissatisfaction_mean(self) -> float:
        """
        Mean dissatisfaction per student.
        """
        return mean(self.supervisor_dissatisfaction_scores())

    def supervisor_dissatisfaction_variance(self) -> float:
        """
        Variance of dissatisfaction scores.
        """
        return variance(self.supervisor_dissatisfaction_scores())

    def write_xlsx(self, filename: str) -> None:
        """
        Writes the solution to an Excel XLSX file (and its problem, for data
        safety).

        Args:
            filename:
                Name of file to write.
        """
        wb = Workbook(write_only=True)  # doesn't create default sheet

        # ---------------------------------------------------------------------
        # Allocations, by student
        # ---------------------------------------------------------------------
        ss = wb.create_sheet(SheetNames.STUDENT_ALLOCATIONS)
        ss.append([
            "Student number",
            "Student name",
            "Project number",
            "Project name",
            "Student's rank of (dissatisfaction with) allocated project",
        ])
        for student, project in self._gen_student_project_pairs():
            ss.append([
                student.number,
                student.name,
                project.number,
                project.name,
                student.dissatisfaction(project),
            ])

        # ---------------------------------------------------------------------
        # Allocations, by project
        # ---------------------------------------------------------------------
        ps = wb.create_sheet(SheetNames.PROJECT_ALLOCATIONS)
        ps.append([
            "Project number",
            "Project name",
            "Student number(s)",
            "Student name(s)",
            "Project supervisor's rank(s) of (dissatisfaction with) allocated student(s)",  # noqa
        ])
        for project in self.problem.sorted_projects():
            student_numbers = []  # type: List[int]
            student_names = []  # type: List[str]
            supervisor_dissatisfactions = []  # type: List[float]
            for student, allocated_proj in self._gen_student_project_pairs():
                if allocated_proj == project:
                    student_numbers.append(student.number)
                    student_names.append(student.name)
                    supervisor_dissatisfactions.append(
                        project.dissatisfaction(student)
                    )
            ps.append([
                project.number,
                project.name,
                ", ".join(str(x) for x in student_numbers),
                ", ".join(student_names),
                ", ".join(str(x) for x in supervisor_dissatisfactions),
            ])

        # ---------------------------------------------------------------------
        # Software and settings information
        # ---------------------------------------------------------------------
        zs = wb.create_sheet(SheetNames.INFORMATION)
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
             1 - self.supervisor_weight],
            ["Overall weight given to supervisor preferences",
             self.supervisor_weight],
            [],
            ["SUMMARY STATISTICS"],
            [],
            ["Student dissatisfaction mean",
             self.student_dissatisfaction_mean()],
            ["Student dissatisfaction variance",
             self.student_dissatisfaction_variance()],
            ["Supervisor dissatisfaction mean",
             self.supervisor_dissatisfaction_mean()],
            ["Supervisor dissatisfaction variance",
             self.supervisor_dissatisfaction_variance()],
        ]
        for row in zs_rows:
            zs.append(row)

        # ---------------------------------------------------------------------
        # Problem definition
        # ---------------------------------------------------------------------
        self.problem.write_to_xlsx_workbook(wb)

        wb.save(filename)

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


# =============================================================================
# Problem
# =============================================================================

class Problem(object):
    """
    Represents the problem (and solves it) -- projects, students.
    """
    def __init__(self,
                 projects: List[Project],
                 students: List[Student]) -> None:
        """
        Args:
            projects:
                List of projects.
            students:
                List of students, with their project preferences.

        Note that the students and projects are put into a "deterministic
        random" order, i.e. deterministically sorted, then shuffled (but with a
        globally fixed random number generator seed). That ensures fairness and
        consistency.
        """
        self.projects = projects
        self.students = students
        # Fix the order:
        self.students.sort()
        random.shuffle(self.students)
        self.projects.sort()
        random.shuffle(self.projects)

    def __str__(self) -> str:
        """
        We re-sort the output for display purposes.
        """
        projects = "\n".join(p.description() for p in self.sorted_projects())
        students = "\n".join(s.description() for s in self.sorted_students())
        return (
            f"Problem:\n"
            f"Projects:\n"
            f"\n"
            f"{projects}\n"
            f"\n"
            f"Students:\n"
            f"\n"
            f"{students}"
        )

    def sorted_students(self) -> List[Student]:
        """
        Students, sorted by number.
        """
        return sorted(self.students, key=lambda s: s.number)

    def sorted_projects(self) -> List[Project]:
        """
        Projects, sorted by number.
        """
        return sorted(self.projects, key=lambda p: p.number)

    def n_students(self) -> int:
        """
        Number of students.
        """
        return len(self.students)

    def n_projects(self) -> int:
        """
        Number of projects.
        """
        return len(self.projects)

    def _make_solution(self,
                       project_indexes: Sequence[int],
                       supervisor_weight: float,
                       validate: bool = True) -> Solution:
        """
        Creates a solution from project index numbers.

        Args:
            project_indexes:
                Indexes (zero-based) of project numbers, one per student,
                in the order of ``self.students``.
            validate:
                validate input? For debugging only.
        """
        if validate:
            n_students = len(self.students)
            assert len(project_indexes) == n_students, (
                "Number of project indices does not match number of students"
            )
        allocation = {}  # type: Dict[Student, Project]
        for student_idx, project_idx in enumerate(project_indexes):
            allocation[self.students[student_idx]] = self.projects[project_idx]
        return Solution(problem=self, allocation=allocation,
                        supervisor_weight=supervisor_weight)

    def best_solution(self,
                      supervisor_weight: float = DEFAULT_SUPERVISOR_WEIGHT,
                      max_time_s: float = DEFAULT_MAX_SECONDS) \
            -> Optional[Solution]:
        """
        Return the best solution.

        Optimize with the MIP package.
        This is extremely impressive.
        See https://python-mip.readthedocs.io/.

        Args:
            supervisor_weight:
                Weight allocated to supervisor preferences; range [0, 1].
                (Student preferences are weighted as 1 minus this.)
            max_time_s:
                Time limit for optimizer (s).
        """
        def varname(s_: int, p_: int) -> str:
            """
            Makes it easier to create/retrieve model variables.
            The indexes are s for student index, p for project index.
            """
            return f"x[{s_},{p_}]"

        assert 0 <= supervisor_weight <= 1
        student_weight = 1 - supervisor_weight
        log.info(
            f"MIP approach giving student preferences weight {student_weight} "
            f"and supervisor preferences weight {supervisor_weight}")
        n_students = len(self.students)
        n_projects = len(self.projects)
        # Dissatisfaction scores for each project
        # CAUTION: get indexes the right way round!
        weighted_dissatisfaction = [
            [
                (
                    student_weight *
                    self.students[s].dissatisfaction(self.projects[p]) +
                    supervisor_weight *
                    self.projects[p].dissatisfaction(self.students[s])
                )
                for p in range(n_projects)  # second index
            ]
            for s in range(n_students)  # first index
        ]  # indexed s, p

        # Model
        m = Model("Student project allocation")
        # Binary variables to optimize, each linking a student to a project
        # CAUTION: get indexes the right way round!
        x = [
            [
                m.add_var(varname(s, p), var_type=BINARY)
                for p in range(n_projects)  # second index
            ]
            for s in range(n_students)  # first index
        ]  # indexed s, p

        # Objective: happy students/supervisors
        m.objective = minimize(xsum(
            x[s][p] * weighted_dissatisfaction[s][p]
            for p in range(n_projects)
            for s in range(n_students)
        ))

        # Constraints
        # - For each student, exactly one project
        for s in range(n_students):
            m += xsum(x[s][p] for p in range(n_projects)) == 1
        # - For each project, up to the maximum number of students
        for p, project in enumerate(self.projects):
            m += xsum(x[s][p] for s in range(n_students)) <= project.max_n_students  # noqa

        # Optimize
        m.optimize(max_seconds=max_time_s)

        # Extract results
        if not m.num_solutions:
            return None
        # for s in range(n_students):
        #     for p in range(n_projects):
        #         log.debug(f"x[{s}][{p}].x = {x[s][p].x}")
        # self._debug_model_vars(m)
        project_indexes = [
            next(p for p in range(n_projects)
                 # if m.var_by_name(varname(s, p)).x >= ALMOST_ONE)
                 if x[s][p].x >= ALMOST_ONE)
            # ... note that the value of a solved variable is var.x
            # If those two expressions are not the same, there's a bug.
            for s in range(n_students)
        ]
        return self._make_solution(project_indexes,
                                   supervisor_weight=supervisor_weight)

    @staticmethod
    def _debug_model_vars(m: Model) -> None:
        """
        Show the names/values of model variables after fitting.
        """
        lines = [f"Variables in model {m.name!r}:"]
        for v in m.vars:
            lines.append(f"{v.name} == {v.x}")
        log.debug("\n".join(lines))

    @classmethod
    def read_data(cls, filename: str) -> "Problem":
        """
        Reads a file, autodetecting its format, and returning the
        :class:`Problem`.
        """
        # File type?
        _, ext = os.path.splitext(filename)
        if ext == EXT_XLSX:
            return cls.read_data_xlsx(filename)
        else:
            raise ValueError(
                f"Don't know how to read file type {ext!r} for {filename!r}")

    @classmethod
    def read_data_xlsx(cls, filename: str) -> "Problem":
        """
        Reads a :class:`Problem` from an Excel XLSX file.
        """
        log.info(f"Reading XLSX file: {filename}")
        wb = load_workbook(filename, read_only=True, keep_vba=False,
                           data_only=True,  keep_links=False)

        # ---------------------------------------------------------------------
        # Projects
        # ---------------------------------------------------------------------
        projects = []  # type: List[Project]
        # These will raise an error if the named sheet does not exist:
        ws_projects = wb[SheetNames.PROJECTS]  # type: Worksheet
        assert (
            ws_projects.cell(row=1, column=1).value ==
            SheetHeadings.PROJECT_NAME and
            ws_projects.cell(row=1, column=2).value ==
            SheetHeadings.MAX_NUMBER_OF_STUDENTS
        ), (
            f"Bad headings to worksheet {SheetNames.PROJECTS}; expected: "
            f"{SheetHeadings.PROJECT_NAME}, "
            f"{SheetHeadings.MAX_NUMBER_OF_STUDENTS:}"
        )
        for row_number, row in enumerate(ws_projects.iter_rows(min_row=2),
                                         start=2):  # type: int, Sequence[Cell]
            project_number = row_number - 1
            project_name = row[0].value
            assert project_name, (
                f"Missing project name in {SheetNames.PROJECTS} "
                f"row {row_number}"
            )
            try:
                max_n_students = int(row[1].value)
            except (ValueError, TypeError):
                raise ValueError(
                    f"Bad max_n_students in {SheetNames.PROJECTS} "
                    f"row {row_number}"
                )
            projects.append(Project(name=project_name,
                                    number=project_number,
                                    max_n_students=max_n_students))
        n_projects = len(projects)
        assert n_projects, "No projects defined!"
        log.info(f"Number of projects: {n_projects}")

        # ---------------------------------------------------------------------
        # Students with their preferences
        # ---------------------------------------------------------------------
        students = []  # type: List[Student]
        ws_students = wb[SheetNames.STUDENT_PREFERENCES]  # type: Worksheet  # noqa
        # Check project headings
        assert all(
            ws_students.cell(row=1, column=i + 2).value == projects[i].name
            for i in range(len(projects))
        ), (
            f"First row of {SheetNames.STUDENT_PREFERENCES} sheet "
            f"must contain all project names in the same order as in the "
            f"{SheetNames.PROJECTS} sheet"
        )
        # Students
        stp_rows = ws_students.iter_rows(min_row=2)
        for row_number, row in enumerate(stp_rows, start=2):
            student_number = row_number - 1
            assert len(row) == n_projects + 1, (
                f"In {SheetNames.STUDENT_PREFERENCES}, student on row "
                f"{student_number + 1} has a preference row of the wrong "
                f"length (expected {n_projects + 1})."
            )
            student_name = row[0].value
            student_preferences = OrderedDict()  # type: Dict[Project, int]
            for project_number, cell in enumerate(row[1:], start=1):
                try:
                    pref = int(cell.value) if cell.value else None
                except (ValueError, TypeError):
                    raise ValueError(
                        f"Bad preference for student {student_name} in "
                        f"{SheetNames.STUDENT_PREFERENCES} "
                        f"row {row_number}")
                project = projects[project_number - 1]
                student_preferences[project] = pref
            students.append(Student(name=student_name,
                                    number=student_number,
                                    preferences=student_preferences,
                                    n_projects=n_projects))
        n_students = len(students)
        log.info(f"Number of students: {n_students}")
        assert n_students >= 1

        # ---------------------------------------------------------------------
        # Supervisor preferences, stored with their project object
        # ---------------------------------------------------------------------
        ws_supervisorprefs = wb[SheetNames.SUPERVISOR_PREFERENCES]  # type: Worksheet  # noqa
        # Accessing cells by (row, column) index is ridiculously slow here, and
        # the time is spent in the internals of openpyxl; specifically, in
        # xml.etree.ElementTree.XMLParser.feed(). That's true even after
        # install lxml as recommended, and specifying the "simple read-only"
        # options. So, it is **much** faster to load all the values like this
        # and then operate on the copies (e.g. ~6 seconds becomes ~1 ms):
        svp_rows = [
            [cell.value for cell in row]
            for row in ws_supervisorprefs.iter_rows()
        ]  # index as : svp_rows[row_zero_based][column_zero_based]

        # Check project headings
        assert all(
            svp_rows[0][i + 1] == projects[i].name
            for i in range(len(projects))
        ), (
            f"First row of {SheetNames.SUPERVISOR_PREFERENCES} sheet "
            f"must contain all project names in the same order as in the "
            f"{SheetNames.PROJECTS} sheet"
        )
        # Check student names
        assert (
            svp_rows[i + 1][0] == students[i].name
            for i in range(len(students))
        ), (
            f"First column of {SheetNames.SUPERVISOR_PREFERENCES} sheet "
            f"must contain all student names in the same order as in the "
            f"{SheetNames.STUDENT_PREFERENCES} sheet"
        )
        for pcol, project in enumerate(projects, start=2):
            supervisor_prefs = OrderedDict()  # type: Dict[Student, int]
            for srow, student in enumerate(students, start=2):
                pref_str = svp_rows[srow - 1][pcol - 1]
                try:
                    pref = int(pref_str) if pref_str else None
                except (ValueError, TypeError):
                    raise ValueError(
                        f"Bad preference at row={srow}, col={pcol} in "
                        f"{SheetNames.SUPERVISOR_PREFERENCES}")
                supervisor_prefs[student] = pref
            project.set_supervisor_preferences(
                n_students=n_students,
                preferences=supervisor_prefs
            )

        # ---------------------------------------------------------------------
        # Create and return the Problem object
        # ---------------------------------------------------------------------
        log.info("... finished reading")
        return Problem(projects=projects, students=students)

    # noinspection DuplicatedCode
    def write_to_xlsx_workbook(self, wb: Workbook) -> None:
        """
        Writes the problem data to a spreadsheet (so it can be saved alongside
        the solution).

        Args:
            wb:
                A :class:`openpyxl.workbook.workbook.Workbook` to which to
                write.
        """
        sorted_projects = self.sorted_projects()
        sorted_students = self.sorted_students()

        project_sheet = wb.create_sheet(SheetNames.PROJECTS)
        project_sheet.append([
            SheetHeadings.PROJECT_NAME,
            SheetHeadings.MAX_NUMBER_OF_STUDENTS
        ])
        for p in sorted_projects:
            project_sheet.append([
                p.name,
                p.max_n_students
            ])

        student_sheet = wb.create_sheet(SheetNames.STUDENT_PREFERENCES)
        student_sheet.append(
            [""] + [p.name for p in sorted_projects]
        )
        for s in sorted_students:
            # noinspection PyTypeChecker
            student_sheet.append(
                [s.name] + [s.preferences.raw_preference(p)
                            for p in sorted_projects]
            )

        supervisor_sheet = wb.create_sheet(SheetNames.SUPERVISOR_PREFERENCES)
        supervisor_sheet.append(
            [""] + [p.name for p in sorted_projects]
        )
        for s in sorted_students:
            # noinspection PyTypeChecker
            supervisor_sheet.append(
                [s.name] + [p.supervisor_preferences.raw_preference(s)
                            for p in sorted_projects]
            )


# =============================================================================
# main
# =============================================================================

def main() -> None:
    """
    Command-line entry point.
    """
    # noinspection PyTypeChecker
    parser = argparse.ArgumentParser(
        formatter_class=RawDescriptionArgumentDefaultsHelpFormatter,
        description=f"""
Allocate students to projects, maximizing some version of happiness.

The input spreadsheet should have the following format (in each case, the
first row is the title row):

    Sheet name:
        {SheetNames.PROJECTS}
    Description:
        List of projects (one per row) and their student capacity.
    Format:
        {SheetHeadings.PROJECT_NAME}    {SheetHeadings.MAX_NUMBER_OF_STUDENTS}
        Project One     1
        Project Two     1
        Project Three   2
        ...             ...
        
    Sheet name:
        {SheetNames.STUDENT_PREFERENCES}
    Description:
        List of students (one per row) and their rank preferences (1 = top, 2 =
        next, etc.) for projects (one per column).
    Format:
        <ignored>       Project One     Project Two     Project Three   ...
        Miss Smith      1               2                               ...
        Mr Jones        2               1               3               ...
        ...             ...             ...             ...             ...
    
    
    Sheet name:
        {SheetNames.SUPERVISOR_PREFERENCES}
    Description:
        List of projects (one per column) and their supervisor's rank
        preferences (1 = top, 2 = next, etc.) for students (one per row).
    Format:
        <ignored>       Project One     Project Two     Project Three   ...
        Miss Smith      1               1                               ...
        Mr Jones        2               2                               ...
        ...             ...             ...

"""  # noqa
    )
    parser.add_argument(
        "filename", type=str,
        help="Spreadsheet filename to read. "
             "Input file types supported: " + str(INPUT_TYPES_SUPPORTED)
    )
    parser.add_argument(
        "--supervisor_weight", type=float, default=DEFAULT_SUPERVISOR_WEIGHT,
        help="Weight allocated to supervisor preferences (student preferences "
             "are weighted as [1 minus this])"
    )
    parser.add_argument(
        "--maxtime", type=float, default=DEFAULT_MAX_SECONDS,
        help="Maximum time (in seconds) to run MIP optimizer for"
    )
    parser.add_argument(
        "--output", type=str,
        help="Optional filename to write output to. "
             "Output types supported: " + str(OUTPUT_TYPES_SUPPORTED)
    )
    parser.add_argument(
        "--verbose", action="store_true",
        help="Be verbose"
    )
    args = parser.parse_args()
    main_only_quicksetup_rootlogger(level=logging.DEBUG if args.verbose
                                    else logging.INFO)

    # Seed RNG
    random.seed(RNG_SEED)

    # Go
    problem = Problem.read_data(args.filename)
    log.info(problem)
    solution = problem.best_solution(
        supervisor_weight=args.supervisor_weight,
        max_time_s=args.maxtime,
    )
    log.info(solution)
    if args.output:
        solution.write_data(args.output)
    else:
        log.warning("Output not saved. Specify the --output option for that.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.critical(str(e))
        raise
