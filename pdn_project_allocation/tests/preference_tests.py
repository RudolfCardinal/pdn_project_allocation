#!/usr/bin/env python

"""
pdn_project_allocation/tests/preferences.py

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

Tests preference functions.

"""

import unittest

from pdn_project_allocation.constants import RankNotation
from pdn_project_allocation.preferences import convert_rank_notation


class PreferenceConversionTests(unittest.TestCase):
    F = RankNotation.FRACTIONAL
    C = RankNotation.COMPETITION
    D = RankNotation.DENSE

    # noinspection DuplicatedCode
    def test_convert_rank_notation(self) -> None:
        groups = (
            # Tuple of lists: fractional, competition, dense.
            ([1.5, 1.5, 3], [1, 1, 3], [1, 1, 2]),
            ([1.5, 1.5, 3.5, 3.5, 5], [1, 1, 3, 3, 5], [1, 1, 2, 2, 3]),
            ([1, 2, 3, 4, 5], [1, 2, 3, 4, 5], [1, 2, 3, 4, 5]),
            ([1, 3, 3, 3, 5], [1, 2, 2, 2, 5], [1, 2, 2, 2, 3]),
            ([1, 3, 3, 3], [1, 2, 2, 2], [1, 2, 2, 2]),
        )
        for f, c, d in groups:
            self.assertEqual(
                convert_rank_notation(f, src=self.F, dst=self.F), f
            )
            self.assertEqual(
                convert_rank_notation(f, src=self.F, dst=self.C), c
            )
            self.assertEqual(
                convert_rank_notation(f, src=self.F, dst=self.D), d
            )

            self.assertEqual(
                convert_rank_notation(c, src=self.C, dst=self.F), f
            )
            self.assertEqual(
                convert_rank_notation(c, src=self.C, dst=self.C), c
            )
            self.assertEqual(
                convert_rank_notation(c, src=self.C, dst=self.D), d
            )

            self.assertEqual(
                convert_rank_notation(d, src=self.D, dst=self.F), f
            )
            self.assertEqual(
                convert_rank_notation(d, src=self.D, dst=self.C), c
            )
            self.assertEqual(
                convert_rank_notation(d, src=self.D, dst=self.D), d
            )

        bad_f = (
            [1, 2, 4],
            [1, 1, 2],
            [1, 1, 3],
            [1.5],
            [2],
            [1, "hello"],
            [1, None],
        )
        for f in bad_f:
            self.assertRaises(ValueError, convert_rank_notation, f, self.F)

        bad_c = (
            [1, 2, 2, 3],
            [1, 1.5, 1.5, 4],
            [2],
            [1.1],
        )
        for c in bad_c:
            self.assertRaises(ValueError, convert_rank_notation, c, self.C)

        bad_d = (
            [1, 1.5, 1.5, 3],
            [1, 1, 3],
            [2],
        )
        for d in bad_d:
            self.assertRaises(ValueError, convert_rank_notation, d, self.D)
