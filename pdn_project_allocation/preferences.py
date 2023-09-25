#!/usr/bin/env python

"""
pdn_project_allocation/preferences.py

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

Preferences class.

"""

from collections import Counter, OrderedDict
import logging
import operator
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from cardinal_pythonlib.maths_py import sum_of_integers_in_inclusive_range

from pdn_project_allocation.constants import (
    DEFAULT_PREFERENCE_POWER,
    DEFAULT_RANK_NOTATION,
    RankNotation,
)

log = logging.getLogger(__name__)


# =============================================================================
# Conversion between ways of expressing preferences, given ties.
# =============================================================================


def convert_rank_notation(
    preferences: Sequence[Union[int, float]],
    src: RankNotation,
    dst: RankNotation = DEFAULT_RANK_NOTATION,
) -> List[Union[int, float]]:
    """
    Converts a list of ranks between rank notation systems.
    Raises ValueError if something is bad.
    """

    def intify(q: Union[float, int]) -> Union[float, int]:
        """
        Convert to int if it is in fact an integer.
        """
        int_q = int(q)
        return int_q if int_q == q else q

    # Input checks
    for x in preferences:
        if not isinstance(x, (int, float)) or x <= 0:
            raise ValueError(f"Bad preferences: {preferences!r}")

    # Count
    n = len(preferences)
    if n == 0:
        return []

    # From source to fractional, stored in "fractional":
    if src == RankNotation.FRACTIONAL:
        # No change
        fractional = list(preferences)
        # ... fulfils promise that result is a list
        # ... also means result not same object as input

    elif src == RankNotation.COMPETITION:
        # Change e.g. 1, 1, 3, 3, 5 -> 1.5, 1.5, 3.5, 3.5, 5
        fractional = []
        prefcount = Counter(preferences)
        for x in preferences:
            if not isinstance(x, int) or not 1 <= x <= n:
                raise ValueError(
                    f"Bad {RankNotation.COMPETITION} preferences: "
                    f"{preferences!r}"
                )
            c = prefcount[x]
            # If there are three "1"s, then they should each have the value
            # (1 + 2 + 3) / 3.
            fractional.append(intify(sum(range(x, x + c)) / c))

    elif src == RankNotation.DENSE:
        # Change e.g. 1, 1, 2, 2, 3 -> 1.5, 1.5, 3.5, 3.5, 5
        fractional = []
        prefcount = Counter(preferences)
        highest_permitted = len(prefcount)
        for x in preferences:
            if not isinstance(x, int) or not 1 <= x <= highest_permitted:
                raise ValueError(
                    f"Bad {RankNotation.DENSE} preferences: {preferences!r}"
                )
            c = prefcount[x]
            n_below = sum(
                rankcount
                for rankval, rankcount in prefcount.items()
                if rankval < x
            )
            fractional.append(
                intify(sum(range(n_below + 1, n_below + 1 + c)) / c)
            )

    else:
        raise AssertionError(f"Bug; bad RankNotation: src = {src!r}")

    # Check intermediate is OK.
    s = sum(fractional)
    t = sum(range(1, n + 1))  # sum from 1 to n inclusive
    if s != t:
        raise ValueError(
            f"Bad preferences {preferences} in {src} notation. In "
            f"{RankNotation.FRACTIONAL} notation that is {fractional}; "
            f"n = {n}, so total should be {t}, but is {s}"
        )

    # From fractional intermediate to destination:
    if dst == RankNotation.FRACTIONAL:
        final = fractional

    elif dst == RankNotation.COMPETITION:
        # Change e.g. 1.5, 1.5, 3.5, 3.5, 5 -> 1, 1, 3, 3, 5
        final = []
        fcount = Counter(fractional)
        for x in fractional:
            refcount = fcount[x]
            lowest = int(x - (refcount - 1) / 2)
            final.append(lowest)

    elif dst == RankNotation.DENSE:
        # Change e.g. 1.5, 1.5, 3.5, 3.5, 5 -> 1, 1, 2, 2, 3
        final = []
        fcount = Counter(fractional)
        for x in fractional:
            n_distinct_values_below = sum(
                1 for rankval in fcount.keys() if rankval < x
            )
            final.append(n_distinct_values_below + 1)

    else:
        raise AssertionError(f"Bad RankNotation: dst = {dst!r}")

    # log.critical(
    #     f"preferences = {preferences}, src = {src}, dst = {dst}, "
    #     f"fractional = {fractional}, final = {final}"
    # )
    return final


# =============================================================================
# Preferences
# =============================================================================


class Preferences(object):
    """
    Represents preference as a mapping from arbitrary objects (being preferred)
    to ranks. Ranks are "dissatisfaction scores", in that the rank of 1 is
    most preferred, and higher-numbered ranks less preferred.
    """

    def __init__(
        self,
        n_options: int,
        preferences: Dict[Any, Union[int, float]] = None,
        owner: Any = None,
        allow_ties: bool = False,
        preference_power: float = DEFAULT_PREFERENCE_POWER,
        input_rank_notation: RankNotation = DEFAULT_RANK_NOTATION,
    ) -> None:
        """
        Args:
            n_options:
                Total number of things to be judged.
            preferences:
                Mapping from "thing being judged" to "rank preference" (1 is
                best). If ``allow_ties`` is set, allows "2.5" for "joint
                second/third"; otherwise, they must be integer.
            owner:
                Person/thing expressing preferences (for cosmetic purposes
                only).
            allow_ties:
                Allows ties to be expressed.
            preference_power:
                Power (exponent) to raise preferences to.
            input_rank_notation:
                Notation for input ranks (see RankNotation).

        Other attributes:
        - ``available_dissatisfaction``: sum of [1 ... ``n_options`]
        - ``allocated_dissatisfaction``: sum of expressed preference ranks.
          (For example, if you only pick your top option, with rank 1, then you
          have expressed a total dissatisfaction of 1. If you have expressed
          a preference for rank #1 and rank #2, you have expressed a total
          dissatisfaction of 3.)

        May raise ValueError if the preferences are bad.
        """
        self._n_options = n_options
        self._preferences = OrderedDict()  # type: Dict[Any, Union[int, float]]
        self._owner = owner
        self._total_dissatisfaction = sum_of_integers_in_inclusive_range(
            1, n_options
        )
        self._allocated_dissatisfaction = 0
        self._allow_ties = allow_ties
        self._preference_power = preference_power

        if not preferences:
            return

        # Convert rank notation to fractional standard.
        # Deal with the fact that some ranks may be None by creating a map
        # from old to new ranks.
        #
        # Don't use filter(None, ...), as that also removes 0 -- which is a
        # dubious value we should catch. Also, it doesn't match the other check
        # below when iterating preferences.items(), so gave KeyError exceptions
        # relating to 0 as a key.
        old_ranks = [v for v in preferences.values() if v is not None]
        new_ranks = convert_rank_notation(
            old_ranks,
            src=input_rank_notation,
            dst=RankNotation.FRACTIONAL,
        )  # May raise ValueError
        rankmap = {old: new for old, new in zip(old_ranks, new_ranks)}
        rankmap[None] = None

        # Store preferences.
        for item, rank in preferences.items():
            if rank is not None:
                self.add(item, rankmap[rank], _validate=False)
                # ... defer validation until all data in...

        self._validate()  # OK, now validate

    def __str__(self) -> str:
        """
        String representation.
        """
        parts = ", ".join(f"{k} → {v}" for k, v in self._preferences.items())
        return (
            f"Preferences({parts}; "
            f"unranked options score {self._unranked_item_dissatisfaction})"
        )

    def __repr__(self) -> str:
        return (
            "{"
            + ", ".join(
                f"{str(k)}: {str(v)}" for k, v in self._preferences.items()
            )
            + "}"
        )

    def set_n_options(self, n_options: int) -> None:
        """
        Sets the total number of options, and ensures that the preferences
        are compatible with this.
        """
        self._n_options = n_options
        self._validate()

    def add(self, item: Any, rank: float, _validate: bool = True) -> None:
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
        if not self._allow_ties:
            assert item not in self._preferences, (
                f"Can't add same item twice (when allow_ties is False); "
                f"attempt to re-add {item!r}"
            )
            assert isinstance(rank, int), (
                f"Only integer preferences allowed "
                f"(when allow_ties is False); was {rank!r}"
            )
            assert rank not in self._preferences.values(), (
                f"No duplicate dissatisfaction scores allowed (when "
                f"allow_ties is False)): attempt to re-add rank {rank}"
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
            assert 1 <= rank <= self._n_options, (
                f"Invalid preference: {rank!r} "
                f"(must be in range [1, {self._n_options}]"
            )
        n_expressed = len(self._preferences)
        expected_allocation = sum_of_integers_in_inclusive_range(
            1, n_expressed
        )
        assert self._allocated_dissatisfaction == expected_allocation, (
            f"For preferences expressed by {self._owner!r}, dissatisfaction "
            f"scores add up to {self._allocated_dissatisfaction}, but must "
            f"add up to {expected_allocation}, since you have expressed "
            f"{n_expressed} preferences (you can only express the 'top n' "
            f"preferences)."
        )
        assert (
            self._allocated_dissatisfaction <= self._total_dissatisfaction
        ), (
            f"Dissatisfaction scores add up to "
            f"{self._allocated_dissatisfaction}, which is more than the "
            f"maximum available of {self._total_dissatisfaction} "
            f"(for {self._n_options} options)"
        )

    @property
    def _unallocated_dissatisfaction(self) -> int:
        """
        The amount of available "dissatisfaction", not yet allocated to an
        item (see :class:`Preferences`).
        """
        return self._total_dissatisfaction - self._allocated_dissatisfaction

    @property
    def _unranked_item_dissatisfaction(self) -> Optional[float]:
        """
        The mean "dissatisfaction" (see :class:`Preferences`) for every option
        without an explicit preference, or ``None`` if there are no such
        options.
        """
        n_unranked = self._n_options - len(self._preferences)
        return (
            self._unallocated_dissatisfaction / n_unranked
            if n_unranked > 0
            else None
        )

    def preference(self, item: Any) -> Union[int, float]:
        """
        Returns a numerical preference score for an item. Will use the
        "unranked" item dissatisfaction if no preference has been expressed for
        this particular item.

        Raises the raw preference score to ``preference_power`` (by default 1).

        Args:
            item:
                The item to look up.
        """
        return self._preferences.get(item, self._unranked_item_dissatisfaction)

    def exponentiated_preference(self, item: Any) -> Union[int, float]:
        """
        As for :meth:`preference`, but raised to ``preference_power`` (by
        default 1).

        Args:
            item:
                The item to look up.
        """
        return self.preference(item) ** self._preference_power

    def raw_preference(self, item: Any) -> Optional[int]:
        """
        Returns the raw preference for an item (for reproducing the input).

        Args:
            item:
                The item to look up.
        """
        return self._preferences.get(item)  # returns None if absent

    def actively_expressed_preference_for(self, item: Any) -> bool:
        """
        Did the person actively express a preference for this item?
        """
        return item in self._preferences

    def items_explicitly_ranked(self) -> List[Any]:
        """
        All the items for which there is an explicit preference.
        """
        return list(self._preferences.keys())

    def items_descending_order(self, all_items: List[Any]) -> List[Any]:
        """
        Returns all the items provided, in descending preference order (or the
        order provided, as a tie-break).
        """
        options = []  # type: List[Tuple[Any, float, int]]
        for i, item in enumerate(all_items):
            preference = self.preference(item)
            options.append((item, preference, i))
        return [
            t[0]  # the item
            for t in sorted(options, key=operator.itemgetter(1, 2))
        ]
        # ... sort by ascending dissatisfaction score (= descending
        # preference), then ascending sequence order

    def is_strict_over(self, items: List[Any]) -> bool:
        """
        Are all preferences strictly ordered for the items in question?
        """
        prefs = [self.preference(item) for item in items]
        n_preferences = len(prefs)
        n_unique_prefs = len(set(prefs))
        return n_preferences == n_unique_prefs

    def is_strict_over_expressed_preferences(self) -> bool:
        """
        Are preferences strictly ordered for the items for which a preference
        has been expressed?
        """
        return self.is_strict_over(self.items_explicitly_ranked())
