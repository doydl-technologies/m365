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
"""Small standard-library ISO/Graph timestamp boundary. Never imports dately."""
import re
from datetime import date, datetime, timezone, timedelta

_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$")


def offset_timezone(value):
    if value.upper() == "Z":
        return timezone.utc
    digits = value[1:].replace(":", "")
    hours, minutes = int(digits[:2]), int(digits[2:])
    if hours > 23 or minutes > 59:
        raise ValueError("Invalid UTC offset")
    return timezone(timedelta(minutes=(60 * hours + minutes) * (-1 if value[0] == "-" else 1)))


def try_iso(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if not isinstance(value, str):
        return None
    value = value.strip()
    # Year and year-month shorthand must not be interpreted as compact MDY.
    if re.fullmatch(r"\d{4}", value):
        return datetime(int(value), 1, 1)
    if re.fullmatch(r"\d{4}-\d{2}", value):
        return datetime(int(value[:4]), int(value[5:]), 1)
    if not _ISO.fullmatch(value):
        return None
    zone = re.search(r"(Z|[+-]\d{2}:?\d{2})$", value) if ":" in value else None
    if zone:
        offset_timezone(zone.group())  # fromisoformat otherwise normalizes bad minutes.
        if len(zone.group()) == 5:
            value = value[:zone.start()] + zone.group()[:3] + ":" + zone.group()[3:]
    # Normalize fractional seconds to Python's microsecond precision;
    # Graph timestamps can contain seven fractional digits.
    value = re.sub(r"\.(\d+)", lambda m: "." + m.group(1)[:6].ljust(6, "0"), value)
    return datetime.fromisoformat(value[:-1] + "+00:00" if value.endswith("Z") else value)


def parse_graph_datetime(value):
    if isinstance(value, str) and not _ISO.fullmatch(value.strip()):
        raise ValueError("Expected a complete ISO Graph date or timestamp")
    parsed = try_iso(value)
    if parsed is None:
        raise ValueError("Expected an ISO Graph timestamp")
    # Graph values are UTC when no offset is supplied; never use the host timezone.
    return parsed.replace(tzinfo=timezone.utc) if parsed.tzinfo is None else parsed
