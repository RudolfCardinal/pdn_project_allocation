.. README.rst

.. _Meld: https://meldmerge.org/

pdn_project_allocation
======================

Allocates projects for the Department of Physiology, Development, and
Neuroscience, University of Cambridge.

By Rudolf Cardinal (rudolf@pobox.com), Department of Psychiatry.


Licence
-------

GNU GPL v3; see https://www.gnu.org/licenses/gpl-3.0.en.html.


Installation
------------

- Create and activate a Python 3 virtual environment.
- Install direct from github:

  .. code-block:: bash

    pip install git+https://github.com/RudolfCardinal/pdn_project_allocation

You should now be able to run the program. Try:

.. code-block:: bash

    pdn_project_allocation --help

To run some automated tests, change into a directory where you're happy to
stash some output files and run

.. code-block:: bash

    pdn_project_allocation_run_tests

This produces solutions to match the test data in the
``pdn_project_allocation/testdata`` directory.


Description of the problem
--------------------------

- There are a number of projects ``p``, and a number of students ``s``.

- Every student needs exactly 1 project.

- Every project can take a certain (project-specific) number of students.

- Students rank projects from 1 (most preferred) onwards.

  - If they don't rank enough, they are treated as being indifferent between
    all other projects.

- Output must be consistent across runs, and consistent against re-ordering of
  students in the input data.

- Supervisors can also express preferences. The overall balance between
  "student satisfaction" and "supervisor satisfaction" is set by a parameter.

- Some student/project combinations can be marked as ineligible (e.g. a student
  might not have the necessary background, no matter how much they want the
  project).


Explanation of how this program works
-------------------------------------

**Absolute constraints**

- The course administrator and the supervisors determine **which projects** are
  available, and **how many students** each project can take.

- The course administrator and the supervisors also determine student
  **eligibility** for each project. (Students might be ineligible because they're
  taking the wrong course modules, or because the supervisor requires that all
  students pre-discuss projects with them and these students haven't, etc.)

**Preferences**

- Students rank their projects: 1 (best), 2 (next best), and so on.

- Supervisors rank potential students: 1 (most preferred), 2 (next), and so on.

- The course administrator decides how much to weight student preferences (e.g.
  70%) versus supervisor preferences (e.g. 30%).

**More on preferences**

- Preferences are expressed via *dissatisfaction scores*. Ranks are naturally
  dissatisfaction scores: if I rank something as #1, and something else as #2,
  I am happiest when I get my lowest score.

- If I am a student and there are 20 projects, I have an allowance of 1 + 2 + 3
  + ... + 20 = 210 dissatisfaction points. My course administrator might allow
  me to rank all 20 projects.

- Alternatively, I might choose to rank only 5 (or the course administrator
  might permit only this). In that case, I will have "used" up scores 1, 2, 3,
  4, and 5 (totalling 15). I will then have 15 unranked projects, and 195
  unallocated dissatisfaction points. In this situation, each of those 15
  projects will be given a dissatisfaction score of 195/15 = 13.

- If the adminstrator permits, I might also express ties. For example, I might
  express "joint second and third" by giving preferences 1, 2.5, 2.5, 4, 5.

- The program enforces the requirement that a student's scores for all the
  projects (those explicitly ranked and those ranked by default) must add up to
  the total dissatisfaction (210 in this example). It also enforces that
  students can only allocate "from the best upwards in rank" -- for example, if
  the student expresses 5 preferences, those scores must add up to 15 (you
  can't say "1, 2, 3, 4, 6").

- Supervisor preferences are handled in exactly the same way. Each project
  supervisor can rank all of the students, or rank some (being indifferent
  between the others), or not rank anyone (having no preference between any
  students). Their dissatisfaction scores are calculated in exactly the same
  way.

If you are trying to express that "the student absolutely cannot do this
project", see *eligibility* above.

If you're the course administrator, consider letting students and supervisors
express as many preferences as they want. It won't cause any harm and may
sometimes help, if competition is fierce for projects.

**Optimization**

- Within hard constraints (every student needs a project; maximum number of
  students per project; eligibility)...

- ... the program maximizes total satisfaction (minimizes total
  dissatisfaction).

  - Specifically, every student-project pairing is associated with
    dissatisfaction from the student, and dissatisfaction from the supervisor,
    as described above. These are weighted (e.g. 70% student, 30% supervisor,
    as above). The total weighted dissatisfaction score is minimized.

- Optimization is achieved via the Python-MIP package
  (https://python-mip.readthedocs.io/), which solves so-called mixed integer
  linear programming problems. This impressive software suite finds optimal
  solutions efficiently.

**Fairness**

- Algorithmic assignment is fair compared to human assignment, in that it
  prevents people "cherry-picking" during manual allocation. It's also fair in
  that it maximizes an objective measure of "happiness" (even though that won't
  exactly reflect real-world happiness).

- It is almost guaranteed, as a reflection of human nature, that students and
  supervisors who didn't get what they wanted will complain about the results
  (or the method). Anticipate this by getting everyone to agree to the
  procedure in advance. Ensure that supervisors are clear about any absolute
  eligibility criteria, convey these to the administrator along with their
  preferences, and agree to accept the result.

- If you run the program several times with the same input, you will get the
  same answers. (It would be unfair otherwise: there would be a temptation to
  keep "flipping the coin" until the operator gets the answer they want.) The
  program achieves this by shuffling its inputs in a "deterministic random" way
  (via a random number generator seed).

- The code is open-source and free for all to use or inspect.

**Advanced options**

- The course administrator may choose to say that students can *only* be
  allocated to projects that they've explicitly ranked. (For example, if a
  student chose 5 most-preferred projects, only those projects can be allocated
  to that student.) However, this may cause the algorithm to fail: there may be
  no such solution. (It is also open to "gaming" if a student is allowed to
  enter only one preference!) If it fails, the program will say so.

- By default, a dissatisfaction score of 2 is "twice as bad" as a score of 1
  (dissatisfaction is linear). Optionally, the course administrator may set
  this to be non-linear by raising dissatisfaction scores to a power
  (exponent). For example, an exponent of 2 would map dissatisfaction scores of
  {1, 2, 3, ...} to {1, 4, 9, ...} for the optimization step.


Changelog
---------

- 2019-10-31: started.

  - Representations.
  - Brute force method.
  - MIP (MILP) method: Mixed Integer Linear Programming Problems.
  - Output.

- 2019-11-01:

  - test framework
  - 1-based dissatisfaction score by default (= rank, probably more
    helpful given that is the input)
  - Failed to find a clear example where you'd be clearly better off with a
    worse mean and a better variance.
  - Experimented with power (exponent); not much gain and adds complexity.

- 2019-11-02:

  - Excel XLSX input/output, in addition to CSV.

- 2019-11-03:

  - Excel only (removed CSV).
  - Supervisors can express preferences too.
  - Removed brute force method; now impractical.
    (With 5 students and 5 projects, one student per project, and no supervisor
    preferences, the brute-force approach examines up to 120 combinations,
    which is fine. With 60 students and 60 projects, then it will examine up to
    8320987112741389895059729406044653910769502602349791711277558941745407315941523456
    = 8.3e81).

- 2020-09-11:

  - Save input data with output.
  - Change default weight to favour students (over supervisors).

- 2020-09-17:

  - Support eligibility.
  - Bugfix to data input checking.

- 2020-09-27, v1.1.0:

  - Option to exponentiate preferences.
  - Configure behaviour for missing eligibility values.
  - Allow projects that permit no students.
  - Show project popularity.
  - Handle Excel sheets that appear to have 1048576 rows (always).
  - Tested with real data.
  - Speed up spreadsheet reading; student CSV output (e.g. for Meld_).

- 2020-09-28, v1.1.1:

  - Shows median/min/max in summary statistics.
  - ``--seed`` option (for debugging ONLY; not fair for real use as it
    encourages fishing for the "right" result from the operator's perspective).
  - Improved README.
  - ``--gs`` option for the Gale-Shapley algorithm, with students as "proposer"
    (https://en.wikipedia.org/wiki/Gale%E2%80%93Shapley_algorithm; see also
    https://www.nrmp.org/nobel-prize/).
