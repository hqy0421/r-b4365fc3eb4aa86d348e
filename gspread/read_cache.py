"""
gspread.read_cache
~~~~~~~~~~~~~~~~~~

In-round caching support for :meth:`gspread.worksheet.Worksheet.batch_get`.

A *read round* is a period of time during which the content of the
worksheets being read is known to be stable.  While such a round is active,
:meth:`~gspread.worksheet.Worksheet.batch_get` remembers which cells have
already been fetched and only asks the API for the cells it has not seen yet.

The cache is intentionally conservative:

* it is scoped by account, spreadsheet, worksheet and by every option that
  changes the shape or the rendering of the returned values;
* a cell is only marked as covered once a successful response has actually
  proven its value (blank cells included);
* failing or truncated responses never mark the unverified part of a range
  as covered.

Coverage is tracked as a set of half-open column intervals per row.  Adjacent
and overlapping intervals are merged immediately, which keeps the bookkeeping
small even when many overlapping A1 ranges are read in a row.
"""

import re
import threading
from bisect import bisect_right
from typing import Any, Dict, Iterable, List, Mapping, NamedTuple, Optional, Tuple

from .utils import (
    SHEET_TITLE_RE,
    DateTimeOption,
    Dimension,
    ValueRenderOption,
    a1_range_to_grid_range,
    absolute_range_name,
    rowcol_to_a1,
)

#: ``(start_row, end_row, start_column, end_column)``, zero-based and half-open.
Bounds = Tuple[int, int, int, int]

_BOUNDS_KEYS = (
    "startRowIndex",
    "endRowIndex",
    "startColumnIndex",
    "endColumnIndex",
)

_SIMPLE_SHEET_TITLE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


class ReadOptions(NamedTuple):
    """Options that change the values returned by a read.

    Two requests can share cached cells only when all of these match.
    """

    major_dimension: str
    value_render_option: str
    date_time_render_option: str

    @classmethod
    def from_request(
        cls,
        major_dimension: Optional[Any] = None,
        value_render_option: Optional[Any] = None,
        date_time_render_option: Optional[Any] = None,
    ) -> "ReadOptions":
        major = _enum_value(major_dimension, Dimension.rows)
        render = _enum_value(value_render_option, ValueRenderOption.formatted)

        if render == str(ValueRenderOption.formatted):
            # ``dateTimeRenderOption`` is ignored by the API in that case, so
            # all variants share the same values.
            date_time = "*"
        else:
            date_time = _enum_value(
                date_time_render_option, DateTimeOption.serial_number
            )

        return cls(major, render, date_time)


class CacheKey(NamedTuple):
    """Full isolation key of one cached worksheet view."""

    account: Tuple[Any, ...]
    spreadsheet_id: str
    worksheet_id: Any
    options: ReadOptions


def _enum_value(value: Optional[Any], default: Any) -> str:
    if value is None:
        value = default
    return str(value)


def a1_part_of_range(range_label: str) -> str:
    """Return the A1 part of a possibly sheet-qualified range label.

    >>> a1_part_of_range("'Sheet 1'!A1:B2")
    'A1:B2'
    >>> a1_part_of_range("Sheet1!A1:B2")
    'A1:B2'
    >>> a1_part_of_range("A1:B2")
    'A1:B2'
    """
    match = SHEET_TITLE_RE.match(range_label)
    if match is not None:
        # The regex consumes the trailing "!" as well.
        return range_label[match.end() :]
    return range_label


def bounds_from_a1(a1_range: str) -> Optional[Bounds]:
    """Parse a bounded A1 range into :data:`Bounds`.

    Sheet-qualified labels are accepted; the sheet prefix is ignored.
    ``None`` is returned for named ranges, unbounded ranges (``A:B``),
    labels that still contain a sheet reference after the prefix is stripped,
    and anything else that cannot be described by a finite rectangle.
    """
    a1_text = a1_part_of_range(a1_range)

    if "!" in a1_text:
        # A range qualified with another sheet cannot be mapped onto this
        # worksheet's cells.
        return None

    try:
        grid_range = a1_range_to_grid_range(a1_text)
    except Exception:
        return None

    if not all(isinstance(grid_range.get(key), int) for key in _BOUNDS_KEYS):
        return None

    start_row = grid_range["startRowIndex"]
    end_row = grid_range["endRowIndex"]
    start_col = grid_range["startColumnIndex"]
    end_col = grid_range["endColumnIndex"]

    if start_row < 0 or start_col < 0 or end_row <= start_row or end_col <= start_col:
        return None

    return (start_row, end_row, start_col, end_col)


def bounds_to_a1(bounds: Bounds) -> str:
    """Format :data:`Bounds` as an A1 range."""
    start_row, end_row, start_col, end_col = bounds

    if end_row - start_row == 1 and end_col - start_col == 1:
        return rowcol_to_a1(start_row + 1, start_col + 1)

    return "{}:{}".format(
        rowcol_to_a1(start_row + 1, start_col + 1),
        rowcol_to_a1(end_row, end_col),
    )


def split_sheet_prefix(range_label: str, sheet_title: str) -> Tuple[str, str]:
    """Split an API range label into its sheet prefix and its A1 part.

    The prefix is returned exactly as the API wrote it (quoted or not) so that
    reconstructed :class:`~gspread.worksheet.ValueRange` objects keep the same
    ``range`` spelling as the API would have used.
    """
    for candidate in (
        absolute_range_name(sheet_title) + "!",
        sheet_title + "!",
    ):
        if range_label.startswith(candidate):
            return candidate, range_label[len(candidate) :]

    match = SHEET_TITLE_RE.match(range_label)
    if match is not None:
        return range_label[: match.end()], range_label[match.end() :]

    if "!" in range_label:
        index = range_label.rfind("!") + 1
        return range_label[:index], range_label[index:]

    return "", range_label


def default_sheet_prefix(sheet_title: str) -> str:
    """Return the sheet prefix the Sheets API is likely to use."""
    if _SIMPLE_SHEET_TITLE_RE.match(sheet_title):
        return sheet_title + "!"
    return absolute_range_name(sheet_title) + "!"


def intersect_bounds(outer: Bounds, inner: Bounds) -> Optional[Bounds]:
    """Return the intersection of two rectangles, or ``None``."""
    start_row = max(outer[0], inner[0])
    end_row = min(outer[1], inner[1])
    start_col = max(outer[2], inner[2])
    end_col = min(outer[3], inner[3])

    if start_row >= end_row or start_col >= end_col:
        return None

    return (start_row, end_row, start_col, end_col)


class Coverage:
    """Compact record of the cells that have been successfully read.

    The data is stored as sorted, disjoint, half-open column intervals per row.
    A full-width read of a 1000 row sheet therefore uses one interval per row,
    while repeated overlapping reads of the same area are merged into a single
    interval instead of accumulating.
    """

    __slots__ = ("_rows",)

    def __init__(self, rows: Optional[Dict[int, List[Tuple[int, int]]]] = None):
        if rows is None:
            self._rows: Dict[int, List[Tuple[int, int]]] = {}
        else:
            self._rows = {row: list(intervals) for row, intervals in rows.items()}

    def copy(self) -> "Coverage":
        return Coverage(self._rows)

    def add(self, bounds: Bounds) -> None:
        """Mark a rectangle as covered, merging with existing intervals."""
        start_row, end_row, start_col, end_col = bounds
        for row in range(start_row, end_row):
            intervals = self._rows.setdefault(row, [])
            _insert_interval(intervals, start_col, end_col)

    def covers(self, row: int, column: int) -> bool:
        intervals = self._rows.get(row)
        if not intervals:
            return False

        index = bisect_right(intervals, (column, float("inf"))) - 1
        if index < 0:
            return False

        start, end = intervals[index]
        return start <= column < end

    def covers_bounds(self, bounds: Bounds) -> bool:
        return not self.missing(bounds)

    def missing(self, bounds: Bounds) -> List[Bounds]:
        """Return the rectangles of ``bounds`` that are not covered yet.

        The returned rectangles never overlap each other and adjacent rows
        covering the same column interval are merged into one rectangle.
        """
        start_row, end_row, start_col, end_col = bounds

        rectangles: List[List[int]] = []
        active: Dict[Tuple[int, int], List[int]] = {}

        for row in range(start_row, end_row):
            runs = _uncovered_runs(self._rows.get(row, []), start_col, end_col)

            if not runs:
                # A fully covered row breaks the vertical continuity: the
                # rectangles above and below must not be merged across it,
                # otherwise covered cells would be requested again.
                active = {}
                continue

            next_active: Dict[Tuple[int, int], List[int]] = {}
            for run in runs:
                rectangle = active.get(run)
                if rectangle is not None:
                    rectangle[1] = row + 1
                else:
                    rectangle = [row, row + 1, run[0], run[1]]
                    rectangles.append(rectangle)
                next_active[run] = rectangle
            active = next_active

        return [(r[0], r[1], r[2], r[3]) for r in rectangles]

    def __len__(self) -> int:
        """Number of stored per-row intervals."""
        return sum(len(intervals) for intervals in self._rows.values())


def _insert_interval(intervals: List[Tuple[int, int]], start: int, end: int) -> None:
    """Insert ``[start, end)`` into a sorted list of disjoint intervals."""
    new_start = start
    new_end = end
    insert_at = 0
    index = 0

    while index < len(intervals):
        current_start, current_end = intervals[index]

        if current_end < new_start:
            insert_at = index + 1
            index += 1
            continue

        if current_start > new_end:
            break

        # Overlapping or adjacent: merge.
        new_start = min(new_start, current_start)
        new_end = max(new_end, current_end)
        intervals.pop(index)

    intervals.insert(insert_at, (new_start, new_end))


def _uncovered_runs(
    intervals: List[Tuple[int, int]], start: int, end: int
) -> List[Tuple[int, int]]:
    """Subtract coverage intervals from ``[start, end)``."""
    runs: List[Tuple[int, int]] = []
    cursor = start

    for covered_start, covered_end in intervals:
        if covered_end <= cursor:
            continue
        if covered_start >= end:
            break

        if covered_start > cursor:
            runs.append((cursor, min(covered_start, end)))

        cursor = max(cursor, covered_end)
        if cursor >= end:
            break

    if cursor < end:
        runs.append((cursor, end))

    return runs


class WorksheetReadCache:
    """Cached values and coverage for one worksheet view."""

    __slots__ = ("_lock", "_coverage", "_values", "_sheet_prefix")

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._coverage = Coverage()
        #: Only non-blank cells are stored; every other covered cell is blank.
        self._values: Dict[Tuple[int, int], Any] = {}
        #: Sheet prefix (``Sheet1!`` or ``'My Sheet'!``) observed in a response.
        self._sheet_prefix: Optional[str] = None

    @property
    def coverage(self) -> Coverage:
        with self._lock:
            return self._coverage.copy()

    def clear(self) -> None:
        with self._lock:
            self._coverage = Coverage()
            self._values.clear()
            self._sheet_prefix = None

    def note_sheet_prefix(self, prefix: str) -> None:
        if not prefix:
            return
        with self._lock:
            if self._sheet_prefix is None:
                self._sheet_prefix = prefix

    def mark_verified(
        self,
        bounds: Bounds,
        values: Mapping[Tuple[int, int], Any],
    ) -> None:
        """Record values that a successful response has proven.

        ``values`` contains the cells the API returned.  Every other cell of
        ``bounds`` is blank as far as this round is concerned: the API omits
        trailing blanks, and a blank answer is still a successful read.  Blank
        cells are not stored individually, which keeps memory proportional to
        the amount of non-empty data while :meth:`snapshot` still returns the
        full rectangle.
        """
        with self._lock:
            for cell, value in values.items():
                if value == "":
                    self._values.pop(cell, None)
                else:
                    self._values[cell] = value

            self._coverage.add(bounds)

    def value(self, row: int, column: int) -> Any:
        with self._lock:
            return self._values.get((row, column), "")

    def snapshot(self, bounds: Bounds) -> List[List[Any]]:
        """Return the raw rectangle of values, blanks included."""
        start_row, end_row, start_col, end_col = bounds
        with self._lock:
            values = self._values
            return [
                [values.get((row, column), "") for column in range(start_col, end_col)]
                for row in range(start_row, end_row)
            ]

    def range_label(self, a1_text: str, sheet_title: str) -> str:
        """Build the ``range`` label for a reconstructed ValueRange.

        ``a1_text`` is kept exactly as the caller spelled it (``A1``,
        ``A1:B2``...) and is prefixed with the sheet spelling observed in an
        earlier response of this round when there is one.
        """
        with self._lock:
            prefix = self._sheet_prefix

        if not prefix:
            prefix = default_sheet_prefix(sheet_title)

        return prefix + a1_text


class ReadCache:
    """Registry of :class:`WorksheetReadCache` entries.

    Entries are isolated by account, spreadsheet, worksheet and render
    options.  :meth:`clear` drops entries; a :class:`ReadCache` instance can
    be passed explicitly to :meth:`Worksheet.batch_get
    <gspread.worksheet.Worksheet.batch_get>` via its ``read_cache`` argument.
    """

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._sheets: Dict[CacheKey, WorksheetReadCache] = {}

    def sheet(
        self,
        account: Tuple[Any, ...],
        spreadsheet_id: str,
        worksheet_id: Any,
        options: ReadOptions,
    ) -> WorksheetReadCache:
        key = CacheKey(account, spreadsheet_id, worksheet_id, options)
        with self._lock:
            cache = self._sheets.get(key)
            if cache is None:
                cache = WorksheetReadCache()
                self._sheets[key] = cache
            return cache

    def clear(
        self,
        account: Optional[Tuple[Any, ...]] = None,
        spreadsheet_id: Optional[str] = None,
        worksheet_id: Any = None,
        clear_worksheet: bool = False,
    ) -> None:
        """Drop cached entries.

        With no argument every entry is dropped.  Filters can be combined to
        target one account, one spreadsheet or one worksheet; all render
        options of the selected worksheet(s) are dropped together.
        """
        with self._lock:
            if account is None and spreadsheet_id is None and not clear_worksheet:
                self._sheets.clear()
                return

            keys = [
                key
                for key in self._sheets
                if (account is None or key.account == account)
                and (spreadsheet_id is None or key.spreadsheet_id == spreadsheet_id)
                and (not clear_worksheet or key.worksheet_id == worksheet_id)
            ]

            for key in keys:
                del self._sheets[key]

    def __len__(self) -> int:
        with self._lock:
            return len(self._sheets)


def account_identity(http_client: Any) -> Tuple[Any, ...]:
    """Build a stable identity for the account behind an HTTP client."""
    auth = getattr(http_client, "auth", None)

    if auth is not None:
        for attribute in (
            "service_account_email",
            "signer_email",
            "client_id",
            "quota_project_id",
        ):
            value = getattr(auth, attribute, None)
            if isinstance(value, str) and value:
                return (type(auth).__name__, attribute, value)

        return (type(auth).__name__, id(auth))

    return ("session", id(getattr(http_client, "session", None)))


def read_round_active(http_client: Any) -> bool:
    """Whether a read round is currently active on ``http_client``."""
    return getattr(http_client, "_read_round_depth", 0) > 0


class ReadRound:
    """Context manager returned by ``read_round()``.

    The cache is emptied when the round starts (so a new round always reads
    fresh values) and dropped again when the round ends.  Rounds can be nested;
    the cache is only dropped when the outermost round ends.
    """

    __slots__ = ("_client", "_entered")

    def __init__(self, client: Any) -> None:
        self._client = client
        self._entered = False

    def __enter__(self) -> "ReadCache":
        self._entered = True
        return self._client.start_read_round()

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        if self._entered:
            self._entered = False
            self._client.end_read_round()


def read_round(client: Any) -> ReadRound:
    """Return a context manager for a read round on ``client``.

    ``client`` may be a :class:`gspread.Client`, a
    :class:`gspread.http_client.HTTPClient`, a :class:`gspread.Spreadsheet`
    or a :class:`gspread.Worksheet`; each of them also exposes the same
    ``read_round()`` method.

    .. versionadded:: 6.3
    """
    return ReadRound(_http_client_of(client))


def _http_client_of(client: Any) -> Any:
    """Find the HTTP client behind a gspread object."""
    http_client = getattr(client, "http_client", None)
    if http_client is not None:
        return http_client

    inner = getattr(client, "client", None)
    if inner is not None and hasattr(inner, "start_read_round"):
        return inner

    if hasattr(client, "start_read_round"):
        return client

    raise TypeError(
        "read_round() expects a gspread Client, Spreadsheet, Worksheet or "
        "HTTPClient, got {!r}".format(type(client).__name__)
    )


class RequestPlan:
    """The requests one cached ``batch_get`` call has to perform.

    ``original_request_indexes`` holds, for every range of the original call,
    the indexes of its pieces in ``requested_ranges``; an empty list means the
    range is served entirely from the cache.
    """

    __slots__ = ("original_request_indexes", "request_bounds", "requested_ranges")

    def __init__(self) -> None:
        self.requested_ranges: List[str] = []
        self.request_bounds: List[Optional[Bounds]] = []
        self.original_request_indexes: List[List[int]] = []

    def add(self, range_label: str, bounds: Optional[Bounds]) -> int:
        index = len(self.requested_ranges)
        self.requested_ranges.append(range_label)
        self.request_bounds.append(bounds)
        return index

    def add_missing(
        self,
        absolute_range: str,
        bounds: Bounds,
        missing: List[Bounds],
        sheet_title: str,
    ) -> List[int]:
        """Schedule the pieces of ``bounds`` that still have to be read.

        The whole range keeps the caller's exact spelling when nothing of it
        is covered yet; otherwise the missing rectangles are requested.
        """
        if not missing:
            return []

        if len(missing) == 1 and missing[0] == bounds:
            return [self.add(absolute_range, bounds)]

        return [
            self.add(
                absolute_range_name(sheet_title, bounds_to_a1(rectangle)),
                rectangle,
            )
            for rectangle in missing
        ]


def is_complete_value_range(entry: Any) -> bool:
    """Whether ``entry`` is a well formed ``ValueRange`` of a batchGet reply."""
    if not isinstance(entry, Mapping):
        return False
    if "range" not in entry or "majorDimension" not in entry:
        return False

    values = entry.get("values")
    if values is None:
        return True

    return isinstance(values, list) and all(isinstance(row, list) for row in values)


def clamp_bounds(bounds: Bounds, grid: Optional[Bounds]) -> Bounds:
    """Intersect ``bounds`` with the worksheet grid, when it is known.

    Cells outside the grid cannot hold values, so caching may treat them as
    covered; the Sheets API itself clips such ranges in its replies.
    """
    if grid is None:
        return bounds

    clamped = intersect_bounds(bounds, grid)
    return bounds if clamped is None else clamped


def prove_bounds(
    entry: Mapping[str, Any],
    requested: Bounds,
    grid: Optional[Bounds] = None,
) -> Optional[Bounds]:
    """Return the region of ``requested`` that a reply has proven.

    Values are only taken from what the reply echoes (see
    :func:`verified_bounds`).  When that echo is exactly the request clipped to
    the worksheet grid, the out-of-grid cells are proven blank as well, so the
    whole rectangle can be covered and never requested again in the round.
    """
    echoed = verified_bounds(entry, requested)
    if echoed is None:
        return None

    if grid is not None and echoed == clamp_bounds(requested, grid):
        return requested

    return echoed


def verified_bounds(entry: Mapping[str, Any], requested: Bounds) -> Optional[Bounds]:
    """Return the part of ``requested`` proven by an API ``ValueRange``.

    A well behaved API echoes the requested range; when it clips a range to
    the grid it echoes a sub-rectangle that starts at the same top-left cell.
    ``None`` is returned for anything else (no echo, a different top-left, a
    larger rectangle), so a misaligned reply can never be cached.
    """
    echoed = bounds_from_a1(a1_part_of_range(str(entry.get("range", ""))))
    if echoed is None:
        return None

    if echoed[0] != requested[0] or echoed[2] != requested[2]:
        return None

    if not bounds_contain(requested, echoed):
        return None

    return echoed


def answers_bounds(entry: Mapping[str, Any], requested: Bounds) -> bool:
    """Whether ``entry`` plausibly answers a request for ``requested``."""
    return verified_bounds(entry, requested) is not None


def bounds_contain(outer: Bounds, inner: Bounds) -> bool:
    """Whether ``inner`` is fully inside ``outer`` (half-open intervals)."""
    return (
        outer[0] <= inner[0]
        and inner[1] <= outer[1]
        and outer[2] <= inner[2]
        and inner[3] <= outer[3]
        and inner[0] < inner[1]
        and inner[2] < inner[3]
    )


def iter_cell_values(
    entry: Mapping[str, Any],
    bounds: Bounds,
    default_major_dimension: str,
) -> Dict[Tuple[int, int], Any]:
    """Map the values of one API ``ValueRange`` onto absolute cell positions.

    ``bounds`` is the rectangle the values are aligned to.  Cells that the API
    trimmed away (trailing blanks) are simply absent from the result; callers
    fill them with blanks.
    """
    values = entry.get("values")
    if not values:
        return {}

    major_dimension = str(entry.get("majorDimension") or default_major_dimension)
    columns_major = major_dimension == str(Dimension.cols)

    cells: Dict[Tuple[int, int], Any] = {}
    for outer_offset, group in enumerate(values):
        if not isinstance(group, Iterable) or isinstance(group, (str, bytes)):
            continue
        for inner_offset, value in enumerate(group):
            cell = _cell_position(bounds, outer_offset, inner_offset, columns_major)
            if cell is not None:
                cells[cell] = value

    return cells


def _cell_position(
    bounds: Bounds,
    outer_offset: int,
    inner_offset: int,
    columns_major: bool,
) -> Optional[Tuple[int, int]]:
    """Translate a value position into an absolute cell, if it is inside."""
    start_row, end_row, start_col, end_col = bounds

    if columns_major:
        row = start_row + inner_offset
        column = start_col + outer_offset
    else:
        row = start_row + outer_offset
        column = start_col + inner_offset

    if start_row <= row < end_row and start_col <= column < end_col:
        return (row, column)

    return None
