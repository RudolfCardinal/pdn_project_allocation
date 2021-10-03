..  pdn_project_allocation/development_notes.rst

..  Copyright (C) 2019-2021 Rudolf Cardinal (rudolf@pobox.com).
    .
    This file is part of pdn_project_allocation.
    .
    This is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    .
    This software is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.
    .
    You should have received a copy of the GNU General Public License
    along with this software. If not, see <http://www.gnu.org/licenses/>.


Development notes
-----------------

Slightly tricky question: optimizing mean versus variance.

- Dissatisfaction mean: lower is better, all else being equal.
- Dissatisfaction variance: lower is better, all else being equal.
- So we have two options:

  - Optimize mean, then use variance as tie-breaker.
  - Optimize a weighted combination of mean and variance.

- Note that "least variance" itself is a rubbish proposition; that can mean
  "consistently bad".

- The choice depends whether greater equality can outweight slightly worse
  mean (dis)satisfaction. I've not found a good example of this. Optimizing
  mean happiness seems to be fine.

- Since moving to a MILP method (see below), we just optimize total weighted
  dissatisfaction.

Consistency:

- For the old brute-force approach, there was a need to break ties randomly,
  e.g. in the case of two projects with both students ranking the first project
  top. Consistency is important and lack of bias (e.g. alphabetical bias) is
  important, so we (a) set a consistent random number seed; (b)
  deterministically and then randomly sort the students; (c) run the optimizer.
  This gives consistent results and does not depend on e.g. alphabetical
  ordering, who comes first in the spreadsheet, etc. (No such effort is applied
  to project ordering.)

- The new MILP method also seems to be consistent.


AIM2007 troubles
----------------

.. code-block:: python

    from matching.games import StudentAllocation

    student_to_preferences = {
        "S1": ["P1", "P2", "P3"],
        "S2": ["P1", "P2", "P3"],
        "S3": ["P4", "P5", "P6"],
        "S4": ["P4", "P5", "P6"],
        "S5": ["P7", "P8", "P9"],
        "S6": ["P7", "P8", "P9"],
        "S7": ["P10", "P1", "P2"],
        "S8": ["P9", "P10", "P1"],
        "S9": ["P8", "P9", "P10"],
        "S10": ["P5", "P6", "P7"],
    }
    supervisor_to_preferences = {
        "SV1": ["S2", "S8", "S1", "S7"],
        "SV2": ["S2", "S1", "S7"],
        "SV3": ["S2", "S1"],
        "SV4": ["S3", "S4"],
        "SV5": ["S3", "S4", "S10"],
        "SV6": ["S3", "S4", "S10"],
        "SV7": ["S5", "S6", "S10"],
        "SV8": ["S5", "S6", "S9"],
        "SV9": ["S8", "S5", "S6", "S9"],
        "SV10": ["S8", "S9", "S7"],
    }
    project_to_supervisor = {
        "P1": "SV1",
        "P2": "SV2",
        "P3": "SV3",
        "P4": "SV4",
        "P5": "SV5",
        "P6": "SV6",
        "P7": "SV7",
        "P8": "SV8",
        "P9": "SV9",
        "P10": "SV10",
    }
    project_to_capacity = {
        "P1": 1,
        "P2": 1,
        "P3": 1,
        "P4": 1,
        "P5": 1,
        "P6": 1,
        "P7": 1,
        "P8": 1,
        "P9": 1,
        "P10": 1,
    }
    supervisor_to_capacity = {
        "SV1": 1,
        "SV2": 1,
        "SV3": 1,
        "SV4": 1,
        "SV5": 1,
        "SV6": 1,
        "SV7": 1,
        "SV8": 1,
        "SV9": 1,
        "SV10": 1,
    }

    game = StudentAllocation.create_from_dictionaries(
        student_to_preferences,
        supervisor_to_preferences,
        project_to_supervisor,
        project_to_capacity,
        supervisor_to_capacity,
    )

    matching = game.solve(optimal="student")
    assert game.check_validity()  # OK
    assert game.check_stability()  # OK

    # But, what it doesn't tell you:

    print(matching)

    # {P1: [S2], P2: [S1], P3: [], P4: [S3], P5: [S4], P6: [S10], P7: [S5], P8: [S6], P9: [S8], P10: [S9]}
    # ... i.e. P3 has no student, and S7 has no project.


Playing with the MIP package
----------------------------

Just for fun, the n-queens problem from
https://python-mip.readthedocs.io/en/latest/examples.html:

.. code-block:: python

    from sys import stdout
    from mip import Model, xsum, MAXIMIZE, BINARY

    # number of queens
    n = 60

    queens = Model()

    x = [[queens.add_var(f"x({i},{j})", var_type=BINARY)
          for j in range(n)] for i in range(n)]

    # one per row
    for i in range(n):
        queens += xsum(x[i][j] for j in range(n)) == 1, f"row({i})"

    # one per column
    for j in range(n):
        queens += xsum(x[i][j] for i in range(n)) == 1, f"col({j})"

    # diagonal \
    for p, k in enumerate(range(2 - n, n - 2 + 1)):
        queens += xsum(x[i][j] for i in range(n) for j in range(n)
                       if i - j == k) <= 1, f"diag1({p})"

    # diagonal /
    for p, k in enumerate(range(3, n + n)):
        queens += xsum(x[i][j] for i in range(n) for j in range(n)
                       if i + j == k) <= 1, f"diag2({p})"

    queens.optimize()

    text = ""
    if queens.num_solutions:
        for i, v in enumerate(queens.vars):
            text += 'Q ' if v.x >= 0.99 else '. '
            if i % n == n-1:
                text += "\n"

    print(text)
    # for v in queens.vars: print(v)
    # for c in queens.constrs: print(c)
    # print(queens.objective)  # blank
