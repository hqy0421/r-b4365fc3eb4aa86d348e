"""Unit tests for the in-round batch read cache.

These tests do not use VCR: they drive ``Worksheet.batch_get`` with a fake
``values_batch_get`` so that requests, coverage and reconstruction can be
asserted precisely.
"""

from unittest.mock import Mock

import pytest

from gspread import utils
from gspread.exceptions import APIError
from gspread.http_client import HTTPClient
from gspread.read_cache import (
    Coverage,
    ReadCache,
    ReadOptions,
    bounds_from_a1,
    bounds_to_a1,
    intersect_bounds,
)
from gspread.worksheet import Worksheet

# A stable 6x6 sheet used by every test.  Cells absent from this map are blank.
SHEET = {
    (0, 0): "A1",
    (0, 1): "B1",
    (0, 3): "D1",
    (1, 1): "B2",
    (3, 0): "A4",
    (3, 1): "B4",
    (3, 3): "D4",
    (5, 5): "F6",
}


def last_non_blank_index(values):
    for index in range(len(values) - 1, -1, -1):
        if values[index] != "":
            return index
    return -1


class FakeBatchGet:
    """Stand-in for ``HTTPClient.values_batch_get``.

    Behaves like the Sheets API: values are trimmed of trailing blanks,
    ranges are clipped to the grid and ``range`` is echoed back with the sheet
    title unquoted.
    """

    def __init__(self, sheet=None, row_count=6, column_count=6):
        self.sheet = dict(SHEET if sheet is None else sheet)
        self.row_count = row_count
        self.column_count = column_count
        self.calls = []
        self.truncate_next_response = False
        self.fail_with = None

    def __call__(self, id, ranges, params=None):
        self.calls.append(list(ranges))

        if self.fail_with is not None:
            raise self.fail_with

        major_dimension = (params or {}).get("majorDimension") or "ROWS"
        value_ranges = [
            self._value_range(range_name, major_dimension) for range_name in ranges
        ]

        if self.truncate_next_response:
            self.truncate_next_response = False
            value_ranges = value_ranges[:1]

        return {"spreadsheetId": id, "valueRanges": value_ranges}

    def _value_range(self, range_name, major_dimension):
        a1 = range_name.split("!")[-1]
        grid = utils.a1_range_to_grid_range(a1)
        start_row = min(grid.get("startRowIndex", 0), self.row_count)
        start_col = min(grid.get("startColumnIndex", 0), self.column_count)
        end_row = min(grid.get("endRowIndex", self.row_count), self.row_count)
        end_col = min(grid.get("endColumnIndex", self.column_count), self.column_count)
        end_row = max(end_row, start_row)
        end_col = max(end_col, start_col)

        if major_dimension == "COLUMNS":
            values = [
                [self.sheet.get((row, column), "") for row in range(start_row, end_row)]
                for column in range(start_col, end_col)
            ]
        else:
            values = [
                [
                    self.sheet.get((row, column), "")
                    for column in range(start_col, end_col)
                ]
                for row in range(start_row, end_row)
            ]

        values = [trim_trailing_blanks(group) for group in values]
        while values and not values[-1]:
            values.pop()

        return {
            "range": "{}!{}".format(
                utils.extract_title_from_range(range_name),
                _clip_range_label(start_row, start_col, end_row, end_col),
            ),
            "majorDimension": major_dimension,
            "values": values,
        }


def trim_trailing_blanks(values):
    length = last_non_blank_index(values)
    return values[: length + 1]


def _clip_range_label(start_row, start_col, end_row, end_col):
    start = utils.rowcol_to_a1(start_row + 1, start_col + 1)
    end = utils.rowcol_to_a1(max(end_row, start_row + 1), max(end_col, start_col + 1))
    if start == end:
        return start
    return "{}:{}".format(start, end)


@pytest.fixture
def batch():
    return FakeBatchGet()


@pytest.fixture
def client(batch):
    client = HTTPClient(auth=None, session=Mock())
    client.values_batch_get = batch
    return client


@pytest.fixture
def worksheet(client):
    properties = {
        "title": "Sheet1",
        "sheetId": 0,
        "gridProperties": {"rowCount": 6, "columnCount": 6},
    }
    return Worksheet(
        spreadsheet=None,
        properties=properties,
        spreadsheet_id="spreadsheet-id",
        client=client,
    )


def flatten(value_ranges):
    return [list(value_range) for value_range in value_ranges]


class TestCoverageTest:
    def test_merge_overlapping_intervals(self):
        coverage = Coverage()
        coverage.add((0, 1, 0, 2))
        coverage.add((0, 1, 1, 4))
        coverage.add((0, 1, 6, 8))

        assert len(coverage) == 2
        assert coverage.covers_bounds((0, 1, 0, 4))
        assert coverage.covers(0, 0)
        assert not coverage.covers(0, 4)

    def test_adjacent_intervals_are_merged(self):
        coverage = Coverage()
        coverage.add((0, 1, 0, 2))
        coverage.add((0, 1, 2, 4))
        assert len(coverage) == 1

    def test_copy_is_independent(self):
        coverage = Coverage()
        coverage.add((0, 1, 0, 2))

        clone = coverage.copy()
        clone.add((0, 1, 2, 4))

        assert len(coverage) == 1
        assert len(clone) == 1
        assert not coverage.covers(0, 2)
        assert clone.covers(0, 2)

    def test_missing_does_not_merge_across_covered_rows(self):
        coverage = Coverage()
        coverage.add((1, 2, 0, 4))  # row 2 is fully covered

        assert coverage.missing((0, 3, 0, 4)) == [(0, 1, 0, 4), (2, 3, 0, 4)]

    def test_missing_merges_identical_column_runs_over_rows(self):
        assert Coverage().missing((0, 3, 0, 2)) == [(0, 3, 0, 2)]

    def test_missing_returns_disjoint_rectangles(self):
        coverage = Coverage()
        coverage.add((1, 2, 1, 3))

        assert coverage.missing((0, 3, 0, 4)) == [
            (0, 1, 0, 4),
            (1, 2, 0, 1),
            (1, 2, 3, 4),
            (2, 3, 0, 4),
        ]

    def test_missing_returns_nothing_when_covered(self):
        coverage = Coverage()
        coverage.add((0, 4, 0, 4))
        assert coverage.missing((1, 3, 1, 3)) == []

    def test_bounds_round_trip(self):
        for a1 in ("A1", "A1:B2", "B4:D4", "AA100:AB200"):
            bounds = bounds_from_a1(a1)
            assert bounds is not None
            assert bounds_to_a1(bounds) == a1

    def test_unbounded_ranges_are_not_cached(self):
        assert bounds_from_a1("A:B") is None
        assert bounds_from_a1("A") is None
        assert bounds_from_a1("named_range") is None

    def test_sheet_prefixed_a1_is_accepted(self):
        assert bounds_from_a1("'Sheet1'!A1:B2") == (0, 2, 0, 2)
        assert bounds_from_a1("Sheet1!A1:B2") == (0, 2, 0, 2)

    def test_intersect(self):
        assert intersect_bounds((0, 4, 0, 4), (2, 6, 2, 6)) == (2, 4, 2, 4)
        assert intersect_bounds((0, 4, 0, 4), (4, 6, 4, 6)) is None


class TestReadOptionsTest:
    def test_defaults_match_api_defaults(self):
        options = ReadOptions.from_request()
        assert options.major_dimension == "ROWS"
        assert options.value_render_option == "FORMATTED_VALUE"

    def test_date_time_is_ignored_for_formatted_values(self):
        a = ReadOptions.from_request(
            value_render_option=utils.ValueRenderOption.formatted,
            date_time_render_option=utils.DateTimeOption.serial_number,
        )
        b = ReadOptions.from_request(
            value_render_option=utils.ValueRenderOption.formatted,
            date_time_render_option=utils.DateTimeOption.formatted_string,
        )
        assert a == b

    def test_date_time_separates_other_render_options(self):
        a = ReadOptions.from_request(
            value_render_option=utils.ValueRenderOption.unformatted,
            date_time_render_option=utils.DateTimeOption.serial_number,
        )
        b = ReadOptions.from_request(
            value_render_option=utils.ValueRenderOption.unformatted,
            date_time_render_option=utils.DateTimeOption.formatted_string,
        )
        assert a != b


class TestBatchGetWithoutRoundTest:
    def test_ordinary_batch_get_is_unchanged(self, worksheet, batch):
        result = worksheet.batch_get(["A1:B1", "B4:D4"])

        assert flatten(result) == [[["A1", "B1"]], [["B4", "", "D4"]]]
        assert [value_range.range for value_range in result] == [
            "Sheet1!A1:B1",
            "Sheet1!B4:D4",
        ]
        assert batch.calls == [["'Sheet1'!A1:B1", "'Sheet1'!B4:D4"]]

        # Exactly what the API returned, trailing blanks included in values.
        assert result[1] == [["B4", "", "D4"]]

    def test_second_identical_call_requests_again(self, worksheet, batch):
        worksheet.batch_get(["A1:B1"])
        worksheet.batch_get(["A1:B1"])

        assert batch.calls == [["'Sheet1'!A1:B1"], ["'Sheet1'!A1:B1"]]

    def test_operation_does_not_start_a_round(self, worksheet):
        assert worksheet.read_cache_active is False
        worksheet.batch_get(["A1:B1"])
        assert worksheet.read_cache_active is False


class TestReadRoundTest:
    def test_full_hit_sends_no_request(self, worksheet, batch):
        with worksheet.read_round():
            first = worksheet.batch_get(["A1:B1"])
            second = worksheet.batch_get(["A1:B1"])

        assert flatten(first) == [[["A1", "B1"]]]
        assert flatten(second) == [[["A1", "B1"]]]
        assert batch.calls == [["'Sheet1'!A1:B1"]]

    def test_only_unread_cells_are_requested(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            worksheet.batch_get(["A1:C2"])

        assert batch.calls == [["'Sheet1'!A1:B1"], ["'Sheet1'!C1", "'Sheet1'!A2:C2"]]

    def test_overlap_within_one_call_is_read_once(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(["A1:C3", "B2:D4"])

        # A1:C3 covers rows 1-3 x A-C; B2:D4 therefore only needs D2:D3 and
        # B4:D4, so no cell is requested twice in this call.
        assert batch.calls == [["'Sheet1'!A1:C3", "'Sheet1'!D2:D3", "'Sheet1'!B4:D4"]]
        assert flatten(result) == [
            [["A1", "B1", ""], ["", "B2", ""], ["", "", ""]],
            [["B2", "", ""], ["", "", ""], ["B4", "", "D4"]],
        ]

    def test_duplicate_requests_are_preserved_in_order(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(["A1:B1", "A1:B1"])

        assert batch.calls == [["'Sheet1'!A1:B1"]]
        assert flatten(result) == [[["A1", "B1"]], [["A1", "B1"]]]

    def test_results_follow_request_order_not_response_order(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(["F6", "A1:B1"])

        assert flatten(result) == [[["F6"]], [["A1", "B1"]]]

    def test_trailing_blanks_are_padded(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(["A1:D4"])

        assert flatten(result) == [
            [
                ["A1", "B1", "", "D1"],
                ["", "B2", "", ""],
                ["", "", "", ""],
                ["A4", "B4", "", "D4"],
            ]
        ]

    def test_range_labels_match_uncached_output(self, worksheet, batch):
        uncached = worksheet.batch_get(["A1:B1", "B4:D4"])
        with worksheet.read_round():
            cached = worksheet.batch_get(["A1:B1", "B4:D4"])

        assert [value_range.range for value_range in cached] == [
            value_range.range for value_range in uncached
        ]
        assert [value_range.major_dimension for value_range in cached] == [
            "ROWS",
            "ROWS",
        ]

    def test_cache_is_per_worksheet(self, client, worksheet, batch):
        other = Worksheet(
            spreadsheet=None,
            properties={
                "title": "Sheet1",
                "sheetId": 42,
                "gridProperties": {"rowCount": 6, "columnCount": 6},
            },
            spreadsheet_id="spreadsheet-id",
            client=client,
        )

        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            other.batch_get(["A1:B1"])

        assert len(batch.calls) == 2

    def test_cache_is_per_spreadsheet(self, client, worksheet, batch):
        other = Worksheet(
            spreadsheet=None,
            properties={
                "title": "Sheet1",
                "sheetId": 0,
                "gridProperties": {"rowCount": 6, "columnCount": 6},
            },
            spreadsheet_id="other-spreadsheet-id",
            client=client,
        )

        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            other.batch_get(["A1:B1"])

        assert len(batch.calls) == 2

    def test_cache_is_per_account(self, batch):
        properties = {
            "title": "Sheet1",
            "sheetId": 0,
            "gridProperties": {"rowCount": 6, "columnCount": 6},
        }
        first_client = HTTPClient(auth=None, session=Mock())
        first_client.values_batch_get = batch
        second_client = HTTPClient(auth=None, session=Mock())
        second_client.values_batch_get = batch

        with first_client.read_round():
            Worksheet(
                spreadsheet=None,
                properties=properties,
                spreadsheet_id="spreadsheet-id",
                client=first_client,
            ).batch_get(["A1:B1"])

        with second_client.read_round():
            Worksheet(
                spreadsheet=None,
                properties=properties,
                spreadsheet_id="spreadsheet-id",
                client=second_client,
            ).batch_get(["A1:B1"])

        assert len(batch.calls) == 2

    def test_cache_is_per_render_options(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            worksheet.batch_get(
                ["A1:B1"], value_render_option=utils.ValueRenderOption.unformatted
            )
            worksheet.batch_get(["A1:B1"], major_dimension=utils.Dimension.cols)

        assert len(batch.calls) == 3

    def test_new_round_reads_again(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])

        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])

        assert batch.calls == [["'Sheet1'!A1:B1"], ["'Sheet1'!A1:B1"]]

    def test_no_requests_on_fully_cached_call(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:D4"])
            request_count = len(batch.calls)

            worksheet.batch_get(["A1:D4"])
            worksheet.batch_get(["B2:C3"])

        assert request_count == 1
        assert len(batch.calls) == 1

    def test_empty_cells_count_as_covered(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(["C1:E3"])
            again = worksheet.batch_get(["C1:E3"])

        assert flatten(result) == [[["", "D1", ""], ["", "", ""], ["", "", ""]]]
        assert flatten(again) == flatten(result)
        assert batch.calls == [["'Sheet1'!C1:E3"]]

    def test_major_dimension_columns(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(
                ["A1:B2"], major_dimension=utils.Dimension.cols
            )
            again = worksheet.batch_get(["A1:B2"], major_dimension=utils.Dimension.cols)

        assert result[0].major_dimension == "COLUMNS"
        assert flatten(result) == [[["A1", ""], ["B1", "B2"]]]
        assert flatten(again) == flatten(result)
        assert batch.calls == [["'Sheet1'!A1:B2"]]

    def test_columns_and_rows_do_not_share_coverage(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B2"])
            worksheet.batch_get(["A1:B2"], major_dimension=utils.Dimension.cols)

        assert len(batch.calls) == 2

    def test_manual_start_stop(self, worksheet, batch):
        worksheet.start_read_round()
        try:
            worksheet.batch_get(["A1:B1"])
            worksheet.batch_get(["A1:B1"])
        finally:
            worksheet.end_read_round()

        assert batch.calls == [["'Sheet1'!A1:B1"]]
        assert worksheet.read_cache_active is False

    def test_end_without_start_is_a_no_op(self, worksheet):
        worksheet.end_read_round()
        assert worksheet.read_cache_active is False

    def test_nested_rounds_keep_cache_until_outermost_ends(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            with worksheet.read_round():
                worksheet.batch_get(["A1:B1"])
            assert worksheet.read_cache_active is True
            worksheet.batch_get(["A1:B1"])

        assert worksheet.read_cache_active is False
        assert batch.calls == [["'Sheet1'!A1:B1"]]

    def test_clear_read_cache_forgets_everything(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            worksheet.clear_read_cache()
            worksheet.batch_get(["A1:B1"])

        assert len(batch.calls) == 2

    def test_unbounded_and_named_ranges_are_never_cached(self, worksheet, batch):
        with worksheet.read_round():
            worksheet.batch_get(["A:B"])
            worksheet.batch_get(["A:B"])

        assert len(batch.calls) == 2


class TestFailureTest:
    def test_api_error_does_not_cache_anything(self, worksheet, batch):
        with worksheet.read_round():
            batch.fail_with = APIError(Mock(ok=False, text="boom", json=lambda: {}))

            with pytest.raises(APIError):
                worksheet.batch_get(["A1:B1"])

            batch.fail_with = None
            result = worksheet.batch_get(["A1:B1"])

        assert flatten(result) == [[["A1", "B1"]]]
        assert batch.calls == [["'Sheet1'!A1:B1"], ["'Sheet1'!A1:B1"]]

    def test_partial_response_only_trusts_verified_prefix(self, worksheet, batch):
        with worksheet.read_round():
            batch.truncate_next_response = True
            # Only the first entry was returned; the second range was never
            # proven, so the result omits it.
            result = worksheet.batch_get(["A1:B1", "D1"])
            assert flatten(result) == [[["A1", "B1"]]]

            # The proven prefix stays cached; only D1 is requested again.
            second = worksheet.batch_get(["A1:B1", "D1"])

        assert flatten(second) == [[["A1", "B1"]], [["D1"]]]
        assert batch.calls == [
            ["'Sheet1'!A1:B1", "'Sheet1'!D1"],
            ["'Sheet1'!D1"],
        ]

    def test_unproven_range_is_not_covered_after_failure(self, worksheet, batch):
        with worksheet.read_round():
            batch.truncate_next_response = True
            worksheet.batch_get(["A1:B1", "D1"])

            # D1 was never confirmed, so it is requested (alone) next time.
            batch.truncate_next_response = False
            worksheet.batch_get(["D1"])

        assert batch.calls[-1] == ["'Sheet1'!D1"]

    def test_partial_response_does_not_corrupt_previous_coverage(
        self, worksheet, batch
    ):
        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            batch.truncate_next_response = True
            worksheet.batch_get(["A1:B1", "C1"])
            worksheet.batch_get(["A1:B1"])

        # A1:B1 was proven before the bad response, so it stays cached; C1
        # was never confirmed and is not cached.
        assert batch.calls == [
            ["'Sheet1'!A1:B1"],
            ["'Sheet1'!C1"],
        ]


class TestCacheParameterTest:
    def test_explicit_cache_registry_is_honoured(self, worksheet, batch):
        cache = ReadCache()

        first = worksheet.batch_get(["A1:B1"], read_cache=cache)
        second = worksheet.batch_get(["A1:B1"], read_cache=cache)

        assert flatten(first) == [[["A1", "B1"]]]
        assert flatten(second) == [[["A1", "B1"]]]
        assert batch.calls == [["'Sheet1'!A1:B1"]]

    def test_explicit_registry_is_isolated_from_client_round(self, worksheet, batch):
        cache = ReadCache()

        with worksheet.read_round():
            worksheet.batch_get(["A1:B1"])
            worksheet.batch_get(["A1:B1"], read_cache=cache)

        assert len(batch.calls) == 2
        assert len(cache) == 1

    def test_cache_alias_is_honoured(self, worksheet, batch):
        cache = ReadCache()

        worksheet.batch_get(["A1:B1"], cache=cache)
        worksheet.batch_get(["A1:B1"], cache=cache)

        assert batch.calls == [["'Sheet1'!A1:B1"]]


class TestClippedResponseTest:
    """The API clips ranges to the grid and echoes the clipped range."""

    def test_range_beyond_the_grid_is_read_once(self, worksheet, batch):
        # The fake sheet is 6x6; A1:H8 extends past the grid on both axes, so
        # only the readable part is requested.
        with worksheet.read_round():
            first = worksheet.batch_get(["A1:H8"])
            second = worksheet.batch_get(["A1:H8"])

        assert flatten(first) == flatten(second)
        assert len(first[0]) == 8
        assert len(first[0][0]) == 8
        assert batch.calls == [["'Sheet1'!A1:F6"]]

    def test_grid_wide_range_beyond_rows_reads_once(self, worksheet, batch):
        with worksheet.read_round():
            first = worksheet.batch_get(["A1:F1000"])
            second = worksheet.batch_get(["C5:C20"])

        assert len(first[0]) == 1000
        assert flatten(second) == [[[""] for _ in range(16)]]
        assert batch.calls == [["'Sheet1'!A1:F6"]]

    def test_out_of_grid_cells_are_blank(self, worksheet, batch):
        with worksheet.read_round():
            result = worksheet.batch_get(["E5:G7"])

        assert flatten(result) == [
            [
                ["", "", ""],
                ["", "F6", ""],
                ["", "", ""],
            ]
        ]
        assert batch.calls == [["'Sheet1'!E5:F6"]]


class TestResponseValidationTest:
    def test_misaligned_response_is_not_cached(self, worksheet, batch):
        original = batch.__call__

        def misaligned(id, ranges, params=None):
            response = original(id, ranges, params)
            # Echo a different top-left cell, as would happen if reply order
            # were scrambled.
            response["valueRanges"][0]["range"] = "Sheet1!F6"
            return response

        with worksheet.read_round():
            worksheet.batch_get = worksheet.batch_get  # keep attribute access
            worksheet.client.values_batch_get = misaligned
            worksheet.batch_get(["A1:B1"])
            worksheet.client.values_batch_get = batch
            worksheet.batch_get(["A1:B1"])

        assert len(batch.calls) == 2

    def test_malformed_entry_makes_the_whole_reply_untrusted(self, worksheet, batch):
        original = batch.__call__

        def malformed_second(id, ranges, params=None):
            response = original(id, ranges, params)
            response["valueRanges"].append("not-a-value-range")
            return response

        with worksheet.read_round():
            worksheet.client.values_batch_get = malformed_second
            first = worksheet.batch_get(["A1:B1", "D1"])
            worksheet.client.values_batch_get = batch
            second = worksheet.batch_get(["A1:B1", "D1"])

        # A reply with a malformed or surplus entry is not trusted at all, so
        # both ranges are requested again.
        assert flatten(first) == []
        assert flatten(second) == [[["A1", "B1"]], [["D1"]]]
        assert batch.calls == [
            ["'Sheet1'!A1:B1", "'Sheet1'!D1"],
            ["'Sheet1'!A1:B1", "'Sheet1'!D1"],
        ]

    def test_surplus_entry_makes_the_whole_reply_untrusted(self, worksheet, batch):
        original = batch.__call__

        def with_extra_entry(id, ranges, params=None):
            response = original(id, ranges, params)
            response["valueRanges"].append(response["valueRanges"][0])
            return response

        with worksheet.read_round():
            worksheet.client.values_batch_get = with_extra_entry
            first = worksheet.batch_get(["A1:B1", "D1"])
            worksheet.client.values_batch_get = batch
            second = worksheet.batch_get(["A1:B1", "D1"])

        assert flatten(first) == []
        assert flatten(second) == [[["A1", "B1"]], [["D1"]]]
        assert len(batch.calls) == 2
