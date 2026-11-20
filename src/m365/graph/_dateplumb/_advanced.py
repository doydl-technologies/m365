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
"""Lazy optional Dately adapter; no parallel general-purpose date parser."""
import importlib
import re
import sys
from copy import deepcopy
from datetime import date, datetime, timedelta
from functools import lru_cache
from .._core_dates import try_iso, offset_timezone

_SUPPORTED = "dately>=4.7.8,<4.8"
_INSTALL = 'Install the optional advanced-date support with: python -m pip install "m365[dates]"'


@lru_cache(maxsize=1)
def _backend():
    try:
        backend = importlib.import_module("dately")
    except ImportError as exc:
        raise ImportError("Advanced dateplumb methods require {}. {}".format(_SUPPORTED, _INSTALL)) from exc
    version = getattr(backend, "__version__", "")
    if not isinstance(version, str) or not re.fullmatch(r"4\.7\.\d+", version) or int(version.rsplit(".", 1)[1]) < 8:
        raise ImportError("Unsupported Dately version {!r}; require {}. {}".format(version, _SUPPORTED, _INSTALL))
    if not all(callable(getattr(backend, key, None)) for key in ("convert_date", "detect_date_format")):
        raise ImportError("Incomplete Dately installation. " + _INSTALL)
    return backend


def _map(value, operation):
    if isinstance(value, list): return [_map(v, operation) for v in value]
    if isinstance(value, tuple): return tuple(_map(v, operation) for v in value)
    np = sys.modules.get("numpy")
    if np is not None and isinstance(value, np.ndarray):
        result = np.empty(value.shape, dtype=object)
        for index in np.ndindex(value.shape): result[index] = operation(value[index])
        return result
    pd = sys.modules.get("pandas")
    if pd is not None and isinstance(value, pd.Series): return value.map(operation)
    return operation(value)


def _value(value, backend):
    if type(value).__name__ == "NaTType": raise ValueError("NaT is not a valid date")
    np = sys.modules.get("numpy")
    if np is not None and isinstance(value, np.datetime64):
        if np.isnat(value): raise ValueError("NaT is not a valid date")
        value = value.astype("datetime64[us]").item()
    parsed = try_iso(value)
    if parsed is None:
        if not isinstance(value, str) or not value.strip():
            raise ValueError("Expected a date string, date, or datetime")
        parsed = backend.convert_date(value)
        if not isinstance(parsed, datetime):
            if isinstance(parsed, date): parsed = datetime.combine(parsed, datetime.min.time())
            else: raise ValueError("Backend did not return a date")
        # Retain a source offset that Dately 4.7.8 otherwise drops and reject
        # an explicit weekday that contradicts the parsed date.
        weekday = re.match(r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Wed|Thu|Fri|Sat|Sun)[, ]", value, re.I)
        if weekday and weekday.group(1).lower() not in (parsed.strftime("%A").lower(), parsed.strftime("%a").lower()):
            raise ValueError("Weekday does not agree with calendar date")
        zone = re.search(r"(Z|[+-]\d{2}:?\d{2})$", value) if re.search(r"\d:\d", value) else None
        if zone: parsed = parsed.replace(tzinfo=offset_timezone(zone.group()))
    return parsed


def _format(value, fmt, backend):
    if not isinstance(fmt, str): raise ValueError("to_format must be a string or None")
    if not fmt: return ""
    if any(code not in "YymdHIMSjUwWV" for code in re.findall(r"%[-#](.)", fmt)):
        raise ValueError("Padding flags require numeric date components")
    # Dately adds timedelta(0) even for formatting, resetting fold. A frozen
    # offset keeps the represented instant's zone fields stable in that case.
    if value.fold and value.tzinfo is not None:
        value = value.replace(tzinfo=offset_timezone(value.strftime("%z")), fold=0)
    return backend.convert_date(value, to_format=fmt)


def _template(value, backend):
    if isinstance(value, (date, datetime)): value = value.isoformat()
    if isinstance(value, str) and re.fullmatch(r"\d{4}", value.strip()): return "%Y"
    if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}", value.strip()): return "%Y-%m"
    return re.sub(r"\s+", " ", backend.detect_date_format(value))


def detect(values):
    backend = _backend()
    def one(value):
        _value(value, backend)  # Reject invalid dates/weekday contradictions.
        return _template(value, backend)
    return _map(values, one)


def conversion(values, to_format=None, delta=0, dict_keys=None, dict_inplace=False):
    backend = _backend()
    if isinstance(delta, bool) or not isinstance(delta, int): raise ValueError("delta must be integer days")
    def convert(value):
        result = _value(value, backend)
        if delta:
            try: result = result + timedelta(days=delta)
            except OverflowError: raise ValueError("Date adjustment is out of range") from None
        return _format(result, to_format, backend) if to_format is not None else result
    def has_dict(value):
        return isinstance(value, dict) or isinstance(value, (list, tuple)) and any(has_dict(v) for v in value)
    if has_dict(values):
        if isinstance(dict_keys, str): dict_keys = [dict_keys]
        if not isinstance(dict_keys, (list, tuple, set)) or not dict_keys or any(not isinstance(k, str) for k in dict_keys):
            raise ValueError("Nonempty dict_keys are required for dictionaries")
        keys = set(dict_keys)
        def visit(value):
            if isinstance(value, dict):
                return {k: (_map(v, convert) if k in keys and not has_dict(v) else visit(v)) for k, v in value.items()}
            if isinstance(value, list): return [visit(v) for v in value]
            if isinstance(value, tuple): return tuple(visit(v) for v in value)
            return value
        result = visit(deepcopy(values))
        if dict_inplace:
            if not isinstance(values, dict): raise ValueError("dict_inplace requires a dictionary")
            values.clear(); values.update(result)
            return None
        return result
    if dict_inplace: raise ValueError("dict_inplace requires a dictionary")
    return _map(values, convert)


def fragment(values, component, ret_format=False):
    backend = _backend()
    component = {"minutes": "minute", "seconds": "second", "microseconds": "microsecond"}.get(component, component)
    if component not in ("year", "month", "day", "weekday", "hour24", "hour12", "minute", "second", "microsecond"):
        raise ValueError("Unknown date component")
    def one(value):
        try:
            parsed = _value(value, backend)
            fmt = _template(value, backend)
        except ValueError:
            return None
        if component in ("hour12", "hour24"):
            result = parsed.strftime("%I" if component == "hour12" else "%H") if re.search(r"%[-#]?[HI]", fmt) else None
        elif component == "month" and ("%b" in fmt or "%B" in fmt):
            result = _format(parsed, "%B" if "%B" in fmt else "%b", backend)
        elif component == "weekday" and ("%a" in fmt or "%A" in fmt):
            result = parsed.strftime("%A" if "%A" in fmt else "%a")
        else:
            # Derive from the already-parsed value and detected template rather
            # than asking the backend to rescan a second time with different
            # cached detection state (notably YYYY-MM's nonexistent day).
            codes = {"year": "Yy", "month": "m", "day": "d", "minute": "M", "second": "S", "microsecond": "f", "weekday": "aA"}
            match = re.search(r"%[-#]?[" + codes[component] + "]", fmt)
            result = _format(parsed, match.group(), backend) if match else None
            if component == "microsecond" and match:
                fraction = re.search(r"\d:\d{2}:\d{2}\.(\d+)", str(value))
                if fraction: result = fraction.group(1)
        return (fmt, result) if ret_format and result is not None else result
    return _map(values, one)


def replace_date(values, year=None, month=None, day=None):
    backend = _backend()
    changes = {}
    for key, value in (("year", year), ("month", month), ("day", day)):
        if value is not None:
            if isinstance(value, bool) or not (isinstance(value, int) or isinstance(value, str) and value.isdigit()):
                raise ValueError("Substitutions must be integers or digit strings")
            changes[key] = int(value)
    def one(value):
        parsed = _value(value, backend)
        updated = parsed.replace(**changes)  # Calendar validation, never rollover.
        if isinstance(value, datetime): return updated
        if isinstance(value, date): return updated.date()
        if not isinstance(value, str): raise ValueError("Expected date string or object")
        source = value.strip()
        suffix = re.search(r"[T ]\d{1,2}:\d{2}.*$", source)
        date_part = source[:suffix.start()] if suffix else source
        fmt = _template(date_part, backend)
        for key, directives in (("year", ("%Y", "%y")), ("month", ("%m", "%-m", "%#m", "%b", "%B")), ("day", ("%d", "%-d", "%#d"))):
            if key in changes and not any(v in fmt for v in directives):
                raise ValueError("Cannot replace a component absent from the input")
        result = _format(updated, fmt, backend)
        if date_part.islower(): result = result.lower()
        elif date_part.isupper(): result = result.upper()
        return result + (suffix.group() if suffix else "")
    return _map(values, one)


def sequence(start_date, end_date, to_format=None):
    backend = _backend()
    start, end = _value(start_date, backend), _value(end_date, backend)
    if (start.tzinfo is None) != (end.tzinfo is None): raise ValueError("Incompatible timezone awareness")
    days = (end - start).days
    if days < 0: return []
    if days >= 100000: raise ValueError("Sequence exceeds 100000-item limit")
    return [_format(start + timedelta(days=i), to_format, backend) if to_format is not None else start + timedelta(days=i)
            for i in range(days + 1)]
