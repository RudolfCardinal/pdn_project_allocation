pdn_project_allocation
======================

Allocates projects for the Department of Physiology, Development, and
Neuroscience, University of Cambridge.

By Rudolf Cardinal (rudolf@pobox.com).


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


Description
-----------

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
  students in the input data. There will be a need to break ties randomly, e.g.
  in the case of two projects with both students ranking the first project top.
  Consistency is important and lack of bias (e.g. alphabetical bias) is
  important, so we (a) set a consistent random number seed; (b)
  deterministically and then randomly sort the students; (c) run the optimizer.
  This gives consistent results and does not depend on e.g. alphabetical
  ordering, who comes first in the spreadsheet, etc. (No such effort is applied
  to project ordering.)

Slightly tricky question: optimizing mean versus variance.

- Dissatisfaction mean: lower is better, all else being equal.
- Dissatisfaction variance: lower is better, all else being equal.
- So we have two options:

  - Optimize mean, then use variance as tie-breaker.
  - Optimize a weighted combination of mean and variance.

- Note that "least variance" itself is a rubbish proposition; that can mean
  "consistently bad".

- The choice depends whether greater equality can outweight slightly worse
  mean (dis)satisfaction.

  - CURRENTLY EXPERIMENTING WITH test4*.csv -- NOT ACHIEVED YET! Optimizing
    mean happiness seems to be fine.


Changelog
---------

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
