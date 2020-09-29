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

**Disadvantages**

- The solution is not always *stable* (the technical meaning of stability is
  given below). However, at present, a supposedly optimal "stable" algorithm
  does not provide projects for all the students with our Sep 2020 real-world
  data (see below), so that's are not much use. The software will report and
  explain any instability.


Methodological considerations ("why not use the Nobel Prize-winning method?")
-----------------------------------------------------------------------------

This is an "assignment problem" or "maximum weighted matching" problem (see
https://en.wikipedia.org/wiki/Assignment_problem).

It is different from the "stable marriage problem" (see
https://en.wikipedia.org/wiki/Stable_marriage_problem), used for
hospital/resident matching in the US via the Gale-Shapley algorithm and
derivatives (https://en.wikipedia.org/wiki/Gale%E2%80%93Shapley_algorithm;
https://www.nrmp.org/nobel-prize/). The stable marriage problem aims to pair
couples (person A and B in each couple) such that there is no pairing A1-B1
where A1 prefers another (B2) over their allocated B1, and B2 *also* prefers A1
to their own allocated partner. That would be unstable, because A1 and B2 could
run away together.

"Maximum satisfaction" problems aren't always stable, and vice versa (see e.g.
Irving et al. 1987, https://doi.org/10.1145/28869.28871, and examples at
https://en.wikipedia.org/wiki/Stable_marriage_problem#Different_stable_matchings).

A supposedly optimal stable algorithm for student-project allocation is that by
Abraham, Irving & Manlove (2007, https://doi.org/10.1016/j.jda.2006.03.006), or
"AIM2007". The "two algorithms" of the title are the one that is
student-optimal, and the one that is supervisor-optimal. These algorithms are
implemented in the Python ``matching`` package
(https://matching.readthedocs.io/). In theory, this also brings extra
sophistication, such as the ability to set supervisor capacity as well as
project capacity. However, that implementation can fail completely (e.g. test
example 4 in the ``testdata`` directory), by failing to allocate some students
to any project. The example has no specific supervisor preferences, ten
projects each with capacity for one student, and preferences like this:

.. code-block:: none

        P1	P2	P3	P4	P5	P6	P7	P8	P9	P10
    S1	1	2	3
    S2	1	2	3
    S3				1	2	3
    S4				1	2	3
    S5							1	2	3
    S6							1	2	3
    S7	2	3								1
    S8	3								1	2
    S9								1	2	3
    S10					1	2	3

The AIM2007 algorithm gave:

.. code-block:: none

    Preferences (re-sorted):

    For student S1, setting preferences: [P1, P2, P3]
    For student S2, setting preferences: [P1, P2, P3]
    For student S3, setting preferences: [P4, P5, P6]
    For student S4, setting preferences: [P4, P5, P6]
    For student S5, setting preferences: [P7, P8, P9]
    For student S6, setting preferences: [P7, P8, P9]
    For student S7, setting preferences: [P10, P1, P2]
    For student S8, setting preferences: [P9, P10, P1]
    For student S9, setting preferences: [P8, P9, P10]
    For student S10, setting preferences: [P5, P6, P7]
    For supervisor Supervisor of P1, setting preferences: [S2, S8, S1, S7]
    For supervisor Supervisor of P2, setting preferences: [S2, S1, S7]
    For supervisor Supervisor of P3, setting preferences: [S2, S1]
    For supervisor Supervisor of P4, setting preferences: [S3, S4]
    For supervisor Supervisor of P5, setting preferences: [S3, S4, S10]
    For supervisor Supervisor of P6, setting preferences: [S3, S4, S10]
    For supervisor Supervisor of P7, setting preferences: [S5, S6, S10]
    For supervisor Supervisor of P8, setting preferences: [S5, S6, S9]
    For supervisor Supervisor of P9, setting preferences: [S8, S5, S6, S9]
    For supervisor Supervisor of P10, setting preferences: [S8, S9, S7]

    Result:

    st  pr  student's rank
    S1	P2  2
    S2	P1  1
    S3	P4  1
    S4	P5  2
    S5	P7  1
    S6	P8  2
    S7	--  --  [projects P1, P2, P10 already taken; P3 free but student didn't want it]
    S8	P9  1
    S9	P10 3
    S10	P6  2

The AIM2007 algorithm requires each supervisor to rank *all* those students
that have ranked *at least one* of their projects
(https://matching.readthedocs.io/en/latest/discussion/student_allocation/index.html#key-definitions).
In the absence of a real ranking, we have to give an arbitrary order.
Nonetheless, in this example, an order was given, across all students who
picked that project, and the algorithm (or this implementation) failed.

In contrast, dissatisfaction minimization solves this happily, e.g. with

.. code-block:: none

    st  pr  student's rank
    S1	P1  1
    S2	P3  3
    S3	P4  1
    S4	P5  2
    S5	P9  3
    S6	P7  1
    S7	P2  3
    S8	P10 2
    S9	P8  1
    S10	P6  2

    ... which is also stable, as it happens.

Likewise, with real data (Sep 2020), large numbers of students were unallocated
by this method.

So: a potential extension for future years is to extend supervisor rankings and
retry an algorithm such as AIM2007, but it can't (apparently) cope with the
current situation.

Another possibility is that the algorithm would have worked if students ranked
more projects. However, that would seem unsatisfactory in the sense that it
would necessarily involve more dissatisfaction to bring stability.

Another possibility is that this is just a known failure mode of AIM2007. For
example, Olaosebikan & Manlove (2020,
https://doi.org/10.1007/s10878-020-00632-x) note that "... exactly the same
students are unassigned in all stable matchings", and their Algorithm 1 has
a termination condition of "until every unassigned student has an empty
preference list" (not that no students are unassigned!).

Since we can't have any student unassigned, and we are now up to Aug 2020 in
the research literature, I shall stop there and offer dissatisfaction
minimization as the best practical option that I've come up with, despite the
fact that it offers some unstable solutions.


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
  - Options to use the AIM2007 algorithms, as above.
