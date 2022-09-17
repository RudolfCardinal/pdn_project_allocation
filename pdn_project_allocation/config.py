#!/usr/bin/env python

"""
pdn_project_allocation/config.py

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

Master config class.

"""

from typing import Any, Dict

from pdn_project_allocation.constants import (
    DEFAULT_MAX_SECONDS,
    DEFAULT_METHOD,
    DEFAULT_PREFERENCE_POWER,
    DEFAULT_SUPERVISOR_WEIGHT,
    OptimizeMethod,
)


# =============================================================================
# Master config
# =============================================================================


class Config(object):
    """
    Master config object.
    """

    def __init__(
        self,
        filename: str,
        allow_defunct_projects: bool = False,
        allow_student_preference_ties: bool = False,
        allow_supervisor_preference_ties: bool = False,
        cmd_args: Dict[str, Any] = None,
        debug_model: bool = False,
        max_time_s: float = DEFAULT_MAX_SECONDS,
        missing_eligibility: bool = None,
        no_shuffle: bool = False,
        optimize_method: OptimizeMethod = DEFAULT_METHOD,
        preference_power: float = DEFAULT_PREFERENCE_POWER,
        student_must_have_choice: bool = False,
        supervisor_weight: float = DEFAULT_SUPERVISOR_WEIGHT,
    ) -> None:
        """
        Reads a file, autodetecting its format, and returning the
        :class:`Problem`.

        Args:
            filename:
                Source data file to read.

            allow_defunct_projects:
                Allow projects that permit no students?
            allow_student_preference_ties:
                Allow students to express preference ties?
            allow_supervisor_preference_ties:
                Allow supervisors to express preference ties?
            cmd_args:
                Copy of command-line arguments
            debug_model:
                Report the MIP model before solving it?
            max_time_s:
                Time limit for MIP optimizer (s).
            missing_eligibility:
                Use ``True`` or ``False`` to treat missing eligibility cells
                as meaning "eligible" or "ineligible", respectively, or
                ``None`` to treat blank cells as invalid.
            no_shuffle:
                Don't shuffle anything. FOR DEBUGGING ONLY.
            optimize_method:
                Method to use for optimizing.
            preference_power:
                Power (exponent) to raise preferences to.
            student_must_have_choice:
                Prevent students being allocated to projects they've not
                explicitly ranked?
            supervisor_weight:
                Weight allocated to supervisor preferences; range [0, 1].
                (Student preferences are weighted as 1 minus this.)
        """
        self.filename = filename

        self.allow_defunct_projects = allow_defunct_projects
        self.allow_student_preference_ties = allow_student_preference_ties
        self.allow_supervisor_preference_ties = (
            allow_supervisor_preference_ties  # noqa
        )
        self.debug_model = debug_model
        self.max_time_s = max_time_s
        self.missing_eligibility = missing_eligibility
        self.no_shuffle = no_shuffle
        self.optimize_method = optimize_method
        self.preference_power = preference_power
        self.student_must_have_choice = student_must_have_choice
        self.supervisor_weight = supervisor_weight

        self.cmd_args = cmd_args

    def __str__(self) -> str:
        return str(self.cmd_args)
