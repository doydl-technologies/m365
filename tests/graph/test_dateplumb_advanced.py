# -*- coding: utf-8 -*-
#
# DOYDL's Microsoft Graph Integration Toolkit — m365
#
# The `m365` package provides reusable Python clients for Microsoft Graph.
# It supports SharePoint document libraries, OneDrive files and folders,
# and Outlook mailboxes through the `m365.graph` namespace.
#
# Applications can search for files and folders, transfer documents, move
# or rename items within a drive, and work with messages and attachments.
# The toolkit separates these operations from tenant-specific authentication,
# configuration, and business logic.
#
# Access is subject to Microsoft Graph capabilities, tenant policies, and
# the permissions granted to the calling application or signed-in user.
# Installing the package does not grant access to Microsoft 365 resources.
#
# Microsoft, Microsoft Graph, Microsoft 365, Outlook, SharePoint, and OneDrive
# are trademarks of Microsoft Corporation. This project is not an official
# Microsoft product and does not imply affiliation, sponsorship, or
# endorsement. Third-party assets remain subject to their own terms and
# notices; the license below does not grant rights to third-party trademarks
# or logos.
#
# Copyright (c) 2026 by DOYDL Technologies. All rights reserved.
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the “Software”), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
"""Compatibility contracts for advanced date handling with the optional backend."""
import copy
import importlib.util
import calendar
import unittest
from datetime import datetime, date, timedelta, timezone, tzinfo

import numpy as np
from m365.graph import _dateplumb as dateplumb


DATELY_AVAILABLE = importlib.util.find_spec("dately") is not None


@unittest.skipUnless(DATELY_AVAILABLE, "optional Dately behavior matrix")
class AdvancedDates(unittest.TestCase):

    def test_all_six_with_supported_backend(self):
        self.assertEqual(dateplumb.pro_parse("2026-11-17"), datetime(2026, 11, 17))
        self.assertEqual(dateplumb.conversion("2026-11-17"), datetime(2026, 11, 17))
        self.assertEqual(dateplumb.detect("2026-11-17"), "%Y-%m-%d")
        self.assertEqual(dateplumb.fragment("2026-11-17", "year"), "2026")
        self.assertEqual(dateplumb.sequence("2026-11-17", "2026-11-18"), [datetime(2026, 11, 17), datetime(2026, 11, 18)])
        self.assertEqual(dateplumb.sub.date("2026-11-17", year=2027), "2027-11-17")

    def test_pro_parse_is_conversion_alias(self):
        self.assertEqual(dateplumb.pro_parse, dateplumb.conversion)
        self.assertEqual(dateplumb.pro_parse("2026-11-17", to_format="%Y/%m/%d", delta=2), "2026/11/19")

    def test_supported_structured_formats(self):
        for text in ("2026-11-17", "2026/11/17", "2026.11.17", "11/17/2026", "17/11/2026", "20261117",
                     "November 17, 2026", "Nov 17 2026", "17 November 2026", "Tuesday, November 17, 2026"):
            with self.subTest(text=text): self.assertEqual(dateplumb.conversion(text), datetime(2026, 11, 17))

    def test_ambiguity_is_fixed_and_not_call_order_dependent(self):
        for previous in ("31/12/2026", "12/31/2026", "2026-11-17"):
            dateplumb.pro_parse(previous)
            self.assertEqual(dateplumb.pro_parse("02/03/2026"), datetime(2026, 2, 3))
            self.assertEqual(dateplumb.detect("02/03/2026"), "%m/%d/%Y")

    def test_partial_dates_and_date_object(self):
        self.assertEqual(dateplumb.conversion("2026-11"), datetime(2026, 11, 1))
        self.assertEqual(dateplumb.conversion("2026"), datetime(2026, 1, 1))
        self.assertEqual(dateplumb.conversion(date(2026, 11, 17)), datetime(2026, 11, 17))
        self.assertIsNone(dateplumb.fragment("2026-11", "day"))

    def test_invalid_dates_and_values_rejected(self):
        for value in ("2026-02-29", "2024-02-30", "2026-13-01", "2026-11-17junk", "next Tuesday", "", None, True, 17):
            with self.subTest(value=value), self.assertRaises(ValueError): dateplumb.conversion(value)
        self.assertEqual(dateplumb.conversion("2024-02-29"), datetime(2024, 2, 29))

    def test_day_delta_rollovers(self):
        self.assertEqual(dateplumb.conversion("2026-12-31", delta=1), datetime(2027, 1, 1))
        self.assertEqual(dateplumb.conversion("2024-03-01", delta=-1), datetime(2024, 2, 29))
        for delta in (True, 1.5, "1"):
            with self.subTest(delta=delta), self.assertRaises(ValueError): dateplumb.conversion("2026-11-17", delta=delta)
        with self.assertRaises(ValueError): dateplumb.conversion("9999-12-31", delta=1)

    def test_timezone_and_fractional_seconds(self):
        z = dateplumb.conversion("2026-11-17T01:02:03.123Z")
        self.assertEqual(z, datetime(2026, 11, 17, 1, 2, 3, 123000, timezone.utc))
        value = dateplumb.conversion("2026-11-17 01:02:03-0530")
        self.assertEqual(value.utcoffset(), timedelta(hours=-5, minutes=-30))
        self.assertEqual(dateplumb.conversion(value).utcoffset(), value.utcoffset())
        for value in ("2026-11-17T24:00:00Z", "2026-11-17T12:60:00Z", "2026-11-17T12:00:00+2460"):
            with self.subTest(value=value), self.assertRaises(ValueError): dateplumb.conversion(value)

    def test_twelve_hour_clock(self):
        for text, hour in (("12:00 AM", 0), ("12:00 PM", 12), ("1:02 PM", 13)):
            self.assertEqual(dateplumb.conversion("2026-11-17 " + text).hour, hour)
        with self.assertRaises(ValueError): dateplumb.conversion("2026-11-17 00:00 AM")

    def test_quick_parser_prefix_integration_no_recursion(self):
        self.assertEqual(dateplumb.parse("2026-11-17 12:34"), datetime(2026, 11, 17, 12, 34))
        self.assertEqual(dateplumb.sub.time("2026-11-17 12:34", hours=13), "2026-11-17 13:34")

    def test_collections_preserve_shapes_and_empty_inputs(self):
        self.assertEqual(dateplumb.conversion([]), [])
        self.assertEqual(dateplumb.conversion(("2026-11-17",)), (datetime(2026, 11, 17),))
        self.assertEqual(dateplumb.conversion([["2026-11-17"]]), [[datetime(2026, 11, 17)]])
        values = np.array([["2026-11-17", "2026-11-18"]])
        original = values.copy()
        output = dateplumb.conversion(values)
        self.assertIsInstance(output, np.ndarray)
        self.assertEqual(output.shape, (1, 2))
        self.assertEqual(output[0, 0], datetime(2026, 11, 17))
        np.testing.assert_array_equal(values, original)
        self.assertEqual(dateplumb.detect(np.empty((0, 2), dtype=str)).shape, (0, 2))
        scalar = dateplumb.conversion(np.array("2026-11-17"))
        self.assertEqual(scalar.shape, ())
        self.assertEqual(scalar.item(), datetime(2026, 11, 17))

    def test_numpy_datetime_and_nat(self):
        self.assertEqual(dateplumb.conversion(np.datetime64("2026-11-17")), datetime(2026, 11, 17))
        with self.assertRaises(ValueError): dateplumb.conversion(np.datetime64("NaT"))

    def test_nested_dictionary_copy_and_selection(self):
        value = {"when": "2026-11-17", "label": "not a date", "children": [{"when": ["2026-11-18"]}]}
        original = copy.deepcopy(value)
        output = dateplumb.conversion(value, dict_keys=["when"], to_format="%d/%m/%Y", delta=1)
        self.assertEqual(value, original)
        self.assertEqual(output["when"], "18/11/2026")
        self.assertEqual(output["children"][0]["when"], ["19/11/2026"])
        self.assertEqual(output["label"], "not a date")

    def test_dictionary_inplace_is_transactional(self):
        value = {"when": "2026-11-17"}
        self.assertIsNone(dateplumb.conversion(value, dict_keys="when", dict_inplace=True))
        self.assertEqual(value["when"], datetime(2026, 11, 17))
        invalid = {"when": "2026-11-17", "nested": {"when": "invalid"}}
        original = copy.deepcopy(invalid)
        with self.assertRaises(ValueError): dateplumb.conversion(invalid, dict_keys=["when"], dict_inplace=True)
        self.assertEqual(invalid, original)

    def test_dictionary_contract_validation(self):
        for keys in (None, [], [1]):
            with self.subTest(keys=keys), self.assertRaises(ValueError): dateplumb.conversion({"date": "2026-11-17"}, dict_keys=keys)
        with self.assertRaises(ValueError): dateplumb.conversion("2026-11-17", dict_inplace=True)
        result = dateplumb.conversion([{"when": "2026-11-17"}], dict_keys="when")
        self.assertEqual(result, [{"when": datetime(2026, 11, 17)}])

    def test_detect_and_portable_padding(self):
        self.assertEqual(dateplumb.detect("2026-1-2"), "%Y-%-m-%-d")
        self.assertEqual(dateplumb.detect("31/12/2026"), "%d/%m/%Y")
        self.assertEqual(dateplumb.detect(["2026-11-17", "Nov 17, 2026"]), ["%Y-%m-%d", "%b %d, %Y"])
        for fmt in ("%Y/%-m/%-d", "%Y/%#m/%#d"):
            self.assertEqual(dateplumb.conversion("2026-01-02", to_format=fmt), "2026/1/2")
        self.assertEqual(dateplumb.conversion("0001-01-02", to_format="%Y-%m-%d"), "0001-01-02")
        for fmt in ("%Q", "%", "%#B"):
            with self.subTest(fmt=fmt), self.assertRaises(ValueError): dateplumb.conversion("2026-11-17", to_format=fmt)

    def test_detect_does_not_accept_invalid_input(self):
        with self.assertRaises(ValueError): dateplumb.detect("2026-02-30")
        with self.assertRaises(ValueError): dateplumb.detect("Monday, November 17, 2026")

    def test_fragment_values_absence_and_format(self):
        value = "2026-11-17T13:04:05.123Z"
        for key, expected in (("year", "2026"), ("month", "11"), ("day", "17"), ("hour24", "13"),
                              ("hour12", "01"), ("minutes", "04"), ("seconds", "05"), ("microseconds", "123")):
            with self.subTest(key=key): self.assertEqual(dateplumb.fragment(value, key), expected)
        self.assertEqual(dateplumb.fragment("2026-11-17", "year", ret_format=True), ("%Y-%m-%d", "2026"))
        self.assertIsNone(dateplumb.fragment("2026-11-17", "hour24"))
        self.assertIsNone(dateplumb.fragment("not a date", "year"))
        self.assertEqual(dateplumb.fragment(["2026-11-17", "2027-01-01"], "year"), ["2026", "2027"])
        with self.assertRaises(ValueError): dateplumb.fragment(value, "invented")

    def test_fragment_month_names_and_weekdays(self):
        self.assertEqual(dateplumb.fragment("Tuesday, November 17, 2026", "weekday"), "Tuesday")
        self.assertEqual(dateplumb.fragment("November 17, 2026", "month"), "November")
        self.assertIsNone(dateplumb.fragment("2026-11-17", "weekday"))

    def test_sequence_inclusive_default_and_formatted(self):
        self.assertEqual(dateplumb.sequence("2024-02-28", "2024-03-01"), [datetime(2024, 2, 28), datetime(2024, 2, 29), datetime(2024, 3, 1)])
        self.assertEqual(dateplumb.sequence("2026-12-31", "2027-01-01", "%Y/%m/%d"), ["2026/12/31", "2027/01/01"])
        self.assertEqual(dateplumb.sequence("2026-11-17", "2026-11-17"), [datetime(2026, 11, 17)])
        self.assertEqual(dateplumb.sequence("2026-11-18", "2026-11-17"), [])

    def test_sequence_timezone_and_time_of_day(self):
        result = dateplumb.sequence("2026-11-17T10:00:00+02:00", "2026-11-18T10:00:00+02:00")
        self.assertEqual(len(result), 2)
        self.assertEqual(result[1].hour, 10)
        self.assertEqual(result[1].utcoffset(), timedelta(hours=2))
        with self.assertRaises(ValueError): dateplumb.sequence("2026-11-17", "2026-11-18T10:00:00Z")
        with self.assertRaises(ValueError): dateplumb.sequence("0001-01-01", "9999-12-31")

    def test_datetime_zone_rules_and_fold_not_replaced_by_fixed_offset(self):
        class RuleZone(tzinfo):
            def utcoffset(self, value): return timedelta(hours=-4 if value.month >= 4 else -5)
            def dst(self, value): return timedelta(0)
        zone = RuleZone()
        start = datetime(2026, 3, 31, 12, tzinfo=zone, fold=1)
        end = datetime(2026, 4, 2, 12, tzinfo=zone)
        converted = dateplumb.conversion(start)
        self.assertIs(converted.tzinfo, zone)
        self.assertEqual(converted.fold, 1)
        result = dateplumb.sequence(start, end)
        self.assertEqual(len(result), 3)
        self.assertIs(result[1].tzinfo, zone)
        self.assertEqual(result[1].utcoffset(), timedelta(hours=-4))
        self.assertEqual(dateplumb.sub.date(start, day=30).fold, 1)

    def test_substitution_passes_all_requested_values(self):
        self.assertEqual(dateplumb.sub.date("2026-11-17", year=2027, month="02", day=3), "2027-02-03")
        self.assertEqual(dateplumb.sub.date("2026-1-2", day=9), "2026-1-9")
        self.assertEqual(dateplumb.sub.date("11/17/2026", year=2027), "11/17/2027")
        self.assertEqual(dateplumb.sub.date("17/11/2026", month=12), "17/12/2026")
        self.assertEqual(dateplumb.sub.date("20261117", year=2027), "20271117")

    def test_substitution_preserves_clock_fraction_and_zone(self):
        source = "2026-11-17T13:04:05.123-05:30"
        self.assertEqual(dateplumb.sub.date(source, day=18), source.replace("11-17T", "11-18T"))
        aware = datetime(2026, 11, 17, 13, tzinfo=timezone.utc)
        self.assertEqual(dateplumb.sub.date(aware, day=18), aware.replace(day=18))
        self.assertEqual(dateplumb.sub.date(date(2026, 11, 17), day=18), date(2026, 11, 18))

    def test_substitution_month_name_and_weekday_remain_valid(self):
        self.assertEqual(dateplumb.sub.date("Tuesday, November 17, 2026", day=18), "Wednesday, November 18, 2026")
        self.assertEqual(dateplumb.sub.date("Nov 17, 2026", month=12), "Dec 17, 2026")
        self.assertEqual(dateplumb.sub.date("nov 17, 2026", month=12), "dec 17, 2026")

    def test_substitution_invalid_calendar_no_rollover(self):
        for changes in ({"year": 2025}, {"day": 30}, {"month": 13}, {"day": True}, {"day": 1.5}):
            with self.subTest(changes=changes), self.assertRaises(ValueError): dateplumb.sub.date("2024-02-29", **changes)
        with self.assertRaises(ValueError): dateplumb.sub.date("2026-11", day=2)
        values = ["2026-11-17", "2026-11-18"]
        self.assertEqual(dateplumb.sub.date(values), values)
        self.assertEqual(dateplumb.sub.date(values, year=2027), ["2027-11-17", "2027-11-18"])

    def test_numpy_all_method_shapes(self):
        values = np.array([["2026-11-17", "2026-11-18"]])
        for output in (dateplumb.detect(values), dateplumb.fragment(values, "day"), dateplumb.sub.date(values, year=2027)):
            self.assertIsInstance(output, np.ndarray)
            self.assertEqual(output.shape, values.shape)

    def test_calendar_boundary_sweep_and_format_roundtrip(self):
        for year in (1, 4, 1900, 2000, 2024, 2026, 9999):
            for month in range(1, 13):
                for day in (1, calendar.monthrange(year, month)[1]):
                    text = "{:04d}-{:02d}-{:02d}".format(year, month, day)
                    with self.subTest(text=text):
                        self.assertEqual(dateplumb.conversion(text), datetime(year, month, day))
                        self.assertEqual(dateplumb.conversion(text, to_format=dateplumb.detect(text)), text)
                        self.assertEqual(dateplumb.sub.date(text, day=1), "{:04d}-{:02d}-01".format(year, month))


@unittest.skipUnless(DATELY_AVAILABLE and importlib.util.find_spec("pandas") is not None, "optional Dately/pandas Series matrix")
class PandasDates(unittest.TestCase):
    def test_series_all_methods_preserve_index_and_name(self):
        import pandas as pd
        values = pd.Series(["2026-11-17", "2026-11-18"], index=[4, 9], name="dates")
        for output in (dateplumb.pro_parse(values), dateplumb.conversion(values), dateplumb.detect(values),
                       dateplumb.fragment(values, "day"), dateplumb.sub.date(values, year=2027)):
            self.assertIsInstance(output, pd.Series)
            self.assertEqual(list(output.index), [4, 9])
            self.assertEqual(output.name, "dates")
        self.assertEqual(values.iloc[0], "2026-11-17")

    def test_nat_rejected_and_timestamp_supported(self):
        import pandas as pd
        self.assertEqual(dateplumb.conversion(pd.Timestamp("2026-11-17")), datetime(2026, 11, 17))
        with self.assertRaises(ValueError): dateplumb.conversion(pd.NaT)


if __name__ == "__main__": unittest.main()
