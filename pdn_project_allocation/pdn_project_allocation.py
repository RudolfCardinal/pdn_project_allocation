#!/usr/bin/env python

"""

Allocates projects for the Department of Physiology, Development, and
Neuroscience, University of Cambridge.

By Rudolf Cardinal.

Description:

- There are a number of projects p, and a number of students s, such that
  p >= s. One student per project (if a project can take >1 student, that
  project gets entered as two projects!).
- Projects are represented by integers from 1...p.
- Students rank projects from 1 (most preferred) upwards, as integers.
- In original task specification, they could rank up to 5 projects, but no
  reason not to extend that. (Indifference between any projects not ranked.)
- If they don't rank enough, they are treated as being indifferent between
  them (meaning that the program will maximize everyone else's satisfaction
  without regard to them).
- Output must be consistent across runs, and consistent against re-ordering of
  students in the input data. There will be random tiebreaks, e.g. in the case
  of two projects with both students ranking the first project top; consistency
  is important and lack of bias is unimportant, so we set a random number seed
  to a fixed value and then allocate randomly. (No such effort is applied to
  project ordering.)

Slightly tricky question 1: optimizing mean versus variance.

- Dissatisfaction mean: lower is better, all else being equal.
- Dissatisfaction variance: lower is better, all else being equal.
- So we have two options:

  - Optimize mean, then use variance as tie-breaker.
  - Optimize a weighted combination of mean and variance.

- Note that "least variance" itself is a rubbish proposition; that can mean
  "consistently bad".

- The choice depends whether greater equality can outweight slightly worse
  mean (dis)satisfaction.

  - CURRENTLY EXPERIMENTING WITH test4*.csv -- NOT ACHIEVED YET!

Changelog:

- 2019-10-31: started.

  - Representations.
  - Brute force method.
  - MIP method: Mixed Integer Linear Programming Problems.
  - Output.

- 2019-11-01:

  - test framework
  - 1-based dissatisfaction score by default (= rank, probably more
    helpful given that is the input)
  - Failed to find a clear example where you'd be clearly better off with a
    worse mean and a better variance.
  - Experimented with power (exponent); not much gain and adds complexity.

"""

import argparse
import csv
from enum import Enum
import itertools
import logging
from math import factorial, inf
import random
from statistics import mean, variance
from typing import Dict, Generator, List, Optional, Sequence, Tuple

from cardinal_pythonlib.logs import main_only_quicksetup_rootlogger
from mip import BINARY, minimize, Model, xsum

log = logging.getLogger(__name__)

ALMOST_ONE = 0.99
# DEFAULT_POWER = 1.0
DEFAULT_MAX_SECONDS = 60
ONE_BASED_DISSATISFACTION_SCORES = True
# ... True means that dissatisfaction scores are basically your rank of your
# allocated project (1 = perfect); False is one lower than that (0 = perfect)
RNG_SEED = 1234  # fixed
VERY_VERBOSE = False  # debugging option


# =============================================================================
# Enum classes
# =============================================================================

class SolveMethod(Enum):
    BRUTE_FORCE = 1
    MIP = 2


# =============================================================================
# Helper functions
# =============================================================================

def sum_of_integers_in_inclusive_range(a: int, b: int) -> int:
    """
    Returns the sum of all integers in the range ``[a, b]``, i.e. from ``a`` to
    ``b`` inclusive.
    
    See
    
    - https://math.stackexchange.com/questions/1842152/finding-the-sum-of-numbers-between-any-two-given-numbers
    """  # noqa
    return int((b - a + 1) * (a + b) / 2)


def n_permutations(n: int, k: int) -> int:
    """
    Returns the number of permutations of length ``k`` from a list of length
    ``n``.

    See https://en.wikipedia.org/wiki/Permutation#k-permutations_of_n.
    """
    assert n > 0 and 0 < k <= n
    return int(factorial(n) / factorial(n - k))


# =============================================================================
# StudentPreferences
# =============================================================================

class Student(object):
    """
    Represents a single student, with their preferences.
    """
    def __init__(self,
                 name: str,
                 number: int,
                 preferences: Dict[int, int],
                 n_projects: int) -> None:
        """
        Args:
            name:
                student's name
            number:
                row number of student (cosmetic only)
            preferences:
                Map from project number (range 1 to n_projects inclusive) to
                dissatisfaction score (range 0 to n_project - 1 inclusive).
            n_projects:
                Total number of projects (for validating inputs).
        """
        self.name = name
        self.number = number
        self.preferences = preferences
        self.n_projects = n_projects

        if ONE_BASED_DISSATISFACTION_SCORES:
            min_dissat = 1
            max_dissat = n_projects
        else:
            min_dissat = 0
            max_dissat = n_projects - 1

        # Precalculate dissatisfaction score for projects not specifically
        # ranked:
        available_dissatisfaction_score = sum_of_integers_in_inclusive_range(
            min_dissat, max_dissat)
        allocated_dissatisfaction_score = sum(self.preferences.values())
        unallocated_dissatisfaction_score = (
            available_dissatisfaction_score - allocated_dissatisfaction_score
        )
        n_prefs = len(preferences)
        n_unranked = n_projects - n_prefs
        self.unranked_dissatisfaction = (
            unallocated_dissatisfaction_score / n_unranked
        ) if n_unranked > 0 else None

        # Validate
        assert all(1 <= pn <= n_projects for pn in preferences.keys()), (
            f"Invalid project number in preferences: {self}"
        )
        prefvalues = list(preferences.values())
        assert all(isinstance(d, int) for d in prefvalues), (
            f"Only integer dissatisfaction score allowed at present: {self}"
        )
        assert all(min_dissat <= d <= max_dissat for d in prefvalues), (
            f"Invalid dissatisfaction score in preferences: {self}"
        )
        assert len(set(prefvalues)) == len(prefvalues), (
            f"No duplicate dissatisfaction scores allowed at present: {self}"
        )
        assert sum(prefvalues) <= available_dissatisfaction_score, (
            f"Dissatisfaction scores add up to more than maximum: {self}"
        )

    def __str__(self) -> str:
        parts = [
            f"P#{k}: {self.preferences[k]}"
            for k in sorted(self.preferences.keys())
        ]
        preferences = ", ".join(parts)
        return (
            f"{self.name} (S#{self.number}): {{{preferences}}} "
            f"(dissatisfaction with unranked projects: "
            f"{self.unranked_dissatisfaction})"
        )

    def shortname(self) -> str:
        """
        Name and number.
        """
        return f"{self.name} (#{self.number})"

    def __lt__(self, other: "Student") -> bool:
        """
        Default sort is by name (case-insensitive).
        """
        return self.name.lower() < other.name.lower()

    def dissatisfaction(self, project_number: int) -> float:
        """
        How dissatisfied is this student if allocated a particular project?

        First choice scores 0; second choice scores 1; etc.
        If the project number isn't in the student's preference list, it
        scores the mean score of all "absent" project number
        """
        return self.preferences.get(project_number,
                                    self.unranked_dissatisfaction)


# =============================================================================
# Project
# =============================================================================

class Project(object):
    """
    Simple representation of a project.
    """
    def __init__(self, name: str, number: int) -> None:
        """
        Args:
            name:
                project name
            number:
                project number
        """
        assert name, "Missing name"
        assert number >= 1, "Bad project number"
        self.name = name
        self.number = number

    def __str__(self) -> str:
        return f"Project #{self.number}: {self.name}"


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
                the :class:`Problem`, defining projects and students
            allocation:
                the mapping of students to projects
        """
        self.problem = problem
        self.allocation = allocation

    def __str__(self) -> str:
        lines = ["Solution:"]
        for student, project in self._gen_student_project_pairs():
            d = student.dissatisfaction(project.number)
            lines.append(f"{student.shortname()} -> "
                         f"{project} (dissatisfaction {d})")
        lines.append("")
        lines.append(f"Dissatisfaction mean: {self.dissatisfaction_mean()}")
        lines.append(f"Dissatisfaction variance: {self.dissatisfaction_variance()}")  # noqa
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
            f", dissatisfaction {self.dissatisfaction_scores()}"
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

    def dissatisfaction_scores(self) -> List[float]:
        """
        All dissatisfaction scores.
        """
        dscores = []  # type: List[float]
        for student in self.problem.students:
            project = self.allocation[student]
            dscores.append(student.dissatisfaction(project.number))
        return dscores

    def dissatisfaction_total(self) -> float:
        """
        Total of dissatisfaction scores.
        """
        return sum(self.dissatisfaction_scores())

    def dissatisfaction_mean(self) -> float:
        """
        Mean dissatisfaction per student.
        """
        return mean(self.dissatisfaction_scores())

    # def dissatisfaction_exponentiated_mean(self, power: float) -> float:
    #     """
    #     Mean of dissatisfaction scores raised to a power.
    #
    #     No longer used.
    #     """
    #     exp_scores = [s ** power for s in self.dissatisfaction_scores()]
    #     return mean(exp_scores)

    def dissatisfaction_variance(self) -> float:
        """
        Variance of dissatisfaction scores.
        """
        return variance(self.dissatisfaction_scores())

    def score(self) -> Tuple[float, float]:
        """
        Score for comparing solutions.
        Used for the brute-force approach.
        """
        return (self.dissatisfaction_mean(),
                self.dissatisfaction_variance())

    @staticmethod
    def worst_possible_score() -> Tuple[float, float]:
        """
        Worst possible score, in the same format as :meth:`score`.
        Used for the brute-force approach.
        """
        return inf, inf

    @staticmethod
    def score_good_enough(score: Tuple[float, float]) -> bool:
        """
        A score that is good enough to stop (e.g. the best possible score),
        in the same format as :meth:`score`.
        Used for the brute-force approach.
        """
        return score[0] <= 0

    def write_csv(self, filename: str) -> None:
        """
        Writes the solution to a CSV file.
        """
        log.info(f"Writing to: {filename}")
        with open(filename, "wt") as f:
            writer = csv.writer(f)
            writer.writerow([
                "Student name",
                "Project number",
                "Project name",
                "Student dissatisfaction with project"
            ])
            for student, project in self._gen_student_project_pairs():
                writer.writerow([
                    student.name,
                    project.number,
                    project.name,
                    student.dissatisfaction(project.number)
                ])


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
                List of projects
            students:
                List of students, with their project preferences.

        Note that the students are put into a "deterministic random" order,
        i.e. deterministically sorted, then shuffled (but with a globally
        fixed random number generator seed).
        """
        self.projects = projects
        self.students = students
        # Fix the order:
        self.students.sort()
        random.shuffle(self.students)

    def __str__(self) -> str:
        """
        We re-sort the output for display purposes.
        """
        projects = "\n".join(str(p) for p in self.projects)
        students = "\n".join(str(s) for s in self.sorted_students())
        return (
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
        Students, sorted by name.
        """
        return sorted(self.students)

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

    def best_solution(self,
                      method: SolveMethod = SolveMethod.MIP,
                      max_time_s: float = DEFAULT_MAX_SECONDS) \
            -> Optional[Solution]:
        """
        Return the best solution.
        """
        if method == SolveMethod.BRUTE_FORCE:
            return self._best_solution_brute_force()
        elif method == SolveMethod.MIP:
            return self._best_solution_mip(max_seconds=max_time_s)
        else:
            raise ValueError(f"Bad solve method: {method!r}")

    def _make_solution(self, project_indexes: Sequence[int],
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
            assert len(set(project_indexes)) == n_students, (
                "Project indices are not unique"
            )
        allocation = {}  # type: Dict[Student, Project]
        for student_idx, project_idx in enumerate(project_indexes):
            allocation[self.students[student_idx]] = self.projects[project_idx]
        return Solution(problem=self, allocation=allocation)

    # -------------------------------------------------------------------------
    # Brute force
    # -------------------------------------------------------------------------

    def _best_solution_brute_force(self) -> Optional[Solution]:
        """
        Brute force method.
        
        Only for playing. With 5 students and 5 projects, this examines 120
        combinations, which is fine. With 60 students and 60 projects, then it
        will examine up to
        8320987112741389895059729406044653910769502602349791711277558941745407315941523456
        = 8.3e81.
        """  # noqa
        log.info("Brute force approach")
        n_expected = n_permutations(self.n_projects(), self.n_students())
        log.info(f"Expecting to test {n_expected} solutions")
        score = Solution.worst_possible_score()
        best = None  # type: Optional[Solution]
        n_tested = 0
        for solution in self._gen_all_solutions():
            if VERY_VERBOSE:
                log.debug(f"Trying: {solution.shortdesc()}")
            n_tested += 1
            s = solution.score()
            if s < score:
                log.debug(f"Improved score from {score} to {s}")
                best = solution
                score = s
                if Solution.score_good_enough(score):
                    log.info("Found good-enough solution; stopping")
                    break
            elif VERY_VERBOSE:
                log.debug(f"Ignoring solution with score {s}")
        log.info(f"Tested {n_tested} solutions")
        return best

    def _gen_all_solutions(self) -> Generator[Solution, None, None]:
        """
        Generates all possible solutions, in a mindless way.
        """
        all_project_indexes = list(range(len(self.projects)))
        n_students = len(self.students)
        for project_indexes in itertools.permutations(all_project_indexes,
                                                      n_students):
            yield self._make_solution(project_indexes)

    # -------------------------------------------------------------------------
    # MIP; see https://python-mip.readthedocs.io/
    # -------------------------------------------------------------------------

    def _best_solution_mip(self,
                           max_seconds: float = DEFAULT_MAX_SECONDS) \
            -> Optional[Solution]:
        """
        Optimize with the MIP package.
        This is extremely impressive.

        Args:
            max_seconds:
                Time limit for optimizer.
        """
        def varname(s_: int, p_: int) -> str:
            """
            Makes it easier to create/retrieve model variables.
            The indexes are s for student index, p for project index.
            """
            return f"x[{s_},{p_}]"

        log.info("MIP approach")
        n_students = len(self.students)
        n_projects = len(self.projects)
        # Student dissatisfaction scores for each project
        # CAUTION: get indexes the right way round!
        dissatisfaction = [
            [
                self.students[s].dissatisfaction(self.projects[p].number)
                for p in range(n_projects)  # second index
            ]
            for s in range(n_students)  # first index
        ]

        # Model
        m = Model("Student project allocation")
        # CAUTION: get indexes the right way round!
        # Binary variables to optimize, each linking a student to a project
        x = [
            [
                m.add_var(varname(s, p), var_type=BINARY)
                for p in range(n_projects)  # second index
            ]
            for s in range(n_students)  # first index
        ]

        # Objective: happy students
        m.objective = minimize(xsum(
            dissatisfaction[s][p] * x[s][p]
            for p in range(n_projects)
            for s in range(n_students)
        ))

        # Constraints
        # - For each student, exactly one project
        for s in range(n_students):
            m += xsum(x[s][p] for p in range(n_projects)) == 1
        # - For each project, zero or one students
        for p in range(n_projects):
            m += xsum(x[s][p] for s in range(n_students)) <= 1

        # Optimize
        m.optimize(max_seconds=max_seconds)

        # Extract results
        if not m.num_solutions:
            return None
        # for s in range(n_students):
        #     for p in range(n_projects):
        #         log.debug(f"x[{s}][{p}].x = {x[s][p].x}")
        self._debug_model_vars(m)
        project_indexes = [
            next(p for p in range(n_projects)
                 # if m.var_by_name(varname(s, p)).x >= ALMOST_ONE)
                 if x[s][p].x >= ALMOST_ONE)
            # ... note that the value of a solved variable is var.x
            # If those two expressions are not the same, there's a bug.
            for s in range(n_students)
        ]
        return self._make_solution(project_indexes)

    @staticmethod
    def _debug_model_vars(m: Model) -> None:
        """
        Show the names/values of model variables after fitting.
        """
        lines = [f"Variables in model {m.name!r}:"]
        for v in m.vars:
            lines.append(f"{v.name} == {v.x}")
        log.debug("\n".join(lines))


# =============================================================================
# Read data in
# =============================================================================

def read_data_csv(filename: str) -> Problem:
    """
    Reads data from a CSV file, giving the problem.

    Format is:

    .. code-block:: none

        ignored,        project_name_1, project_name_2, ...
        student_name_1, rank_or_blank,  rank_or_blank, ...
        student_name_2, rank_or_blank,  rank_or_blank, ...
        ...

    """
    projects = []  # type: List[Project]
    students = []  # type: List[Student]
    log.info(f"Reading file: {filename}")
    with open(filename, "rt") as f:
        reader = csv.reader(f)

        # First row: read projects
        firstrow = next(reader)
        n_projects = len(firstrow) - 1
        log.info(f"Number of projects: {n_projects}")
        assert n_projects >= 1
        for pn in range(1, n_projects + 1):
            projects.append(Project(name=firstrow[pn], number=pn))

        # Other rows: students and preferences
        for student_number, row in enumerate(reader, start=1):
            student_name = row[0]
            prefs = {}  # type: Dict[int, int]
            for pn in range(1, n_projects + 1):
                rank_str = row[pn]
                if rank_str:
                    try:
                        dissatisfaction_score = int(rank_str)
                        if not ONE_BASED_DISSATISFACTION_SCORES:
                            dissatisfaction_score = dissatisfaction_score - 1
                        prefs[pn] = dissatisfaction_score
                    except (ValueError, TypeError):
                        raise ValueError(f"Bad preference: {rank_str!r}")
            students.append(Student(name=student_name,
                                    number=student_number,
                                    preferences=prefs,
                                    n_projects=n_projects))

    n_students = len(students)
    log.info(f"Number of students: {n_students}")
    assert n_students >= 1
    return Problem(projects=projects, students=students)


# =============================================================================
# main
# =============================================================================

def main() -> None:
    """
    Command-line entry point.
    """
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument(
        "filename", type=str,
        help="CSV filename to read. Top left cell is ignored. "
             "First row (starting with second cell) contains project names. "
             "Other rows are one line per student; "
             "first column contains student names; "
             "other columns contain project-specific ranks "
             "(1 best, 2 second, etc.)"
    )
    parser.add_argument(
        "--maxtime", type=float, default=DEFAULT_MAX_SECONDS,
        help="Maximum time (in seconds) to run MIP optimizer for"
    )
    # parser.add_argument(
    #     "--power", type=float, default=DEFAULT_POWER,
    #     help="We optimize dissatisfaction ^ power. What power should we use?"
    # )
    parser.add_argument(
        "--bruteforce", action="store_true",
        help="Use brute-force method (only for debugging!)"
    )
    parser.add_argument(
        "--output", type=str,
        help="Optional filename to write output to"
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
    problem = read_data_csv(args.filename)
    log.info(f"Problem:\n{problem}")
    solution = problem.best_solution(
        method=SolveMethod.BRUTE_FORCE if args.bruteforce else SolveMethod.MIP,
        # power=args.power,
        max_time_s=args.maxtime,
    )
    log.info(solution)
    if args.output:
        solution.write_csv(args.output)
    else:
        log.warning("Output not saved. Specify the --output option for that.")


if __name__ == "__main__":
    main()
