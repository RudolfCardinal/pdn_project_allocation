#!/usr/bin/env python

"""
pdn_project_allocation/supervisor.py

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

Supervisor class.

"""

from cardinal_pythonlib.reprfunc import auto_repr


# =============================================================================
# Supervisor
# =============================================================================


class Supervisor:
    """
    Simple representation of a supervisor.
    """

    def __init__(
        self,
        name: str,
        number: int,
        max_n_projects: int = None,
        max_n_students: int = None,
    ) -> None:
        """
        Args:
            name:
                Supervisor name.
            number:
                Supervisor number (cosmetic only: matches input order).
            max_n_projects:
                Maximum number of projects this supervisor can supervise.
                (They may offer more projects, but be unable to support all of
                them simultaneously.)
            max_n_students:
                Maximum number of students this supervisor can take.
        """
        assert name, "Missing supervisor name"
        assert number >= 1, "Bad supervisor number"
        assert max_n_projects is None or max_n_projects >= 1, (
            f"Supervisor {name!r}: invalid max_n_projects; must be None or "
            f">=1 but is {max_n_projects!r}"
        )
        assert max_n_students is None or max_n_students >= 1, (
            f"Supervisor {name!r}: invalid max_n_students; must be None or "
            f">=1 but is {max_n_projects!r}"
        )
        self.name = name
        self.number = number
        self.max_n_projects = max_n_projects
        self.max_n_students = max_n_students

    def __str__(self) -> str:
        """
        String representation.
        """
        return f"{self.name} (Sv#{self.number})"

    def __repr__(self) -> str:
        return auto_repr(self)

    def description(self) -> str:
        """
        Verbose description.
        """
        return (
            f"{self}: max_n_projects={self.max_n_projects}, "
            f"max_n_students={self.max_n_students}"
        )
