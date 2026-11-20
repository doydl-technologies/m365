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
from datetime import datetime as dt, timedelta as td, date as d, timezone as tmz, time as t
import re
import time
import sys

from .proxy import Proxy
from . import _advanced
from .._core_dates import try_iso, offset_timezone

# Resolve optional modules only when an operation first uses them.
pd = Proxy("pandas", None)
np = Proxy("numpy", None)

# Basic Graph/ISO parsing is local. Advanced conversion, detection and
# component operations delegate to the optional Dately adapter.

class dt_parse:
    def __init__(self):
        self.date_formats = [
            '%Y.%m.%d', 
            '%m.%d.%Y', 
            '%B %d, %Y', 
            '%b %d, %Y', 
            '%d.%m.%Y', 
            '%d %b %Y', 
            '%d %B %Y',
            '%b %d %Y',
            '%Y%m%d',
            '%d%m%Y',
            '%A, %B %d %Y',
            '%Y-%m-%dT%H:%M:%SZ',
            '%a, %d %b %Y %H: %M:%S',
            '%Y-%m',
            '%Y-%m-%d %H:%M',
            '%Y-%m-%d %I:%M %p',           
        ]
        self.iso_format = re.compile(
            r'^\d{4}-\d{2}-\d{2}'             # Date: YYYY-MM-DD
            r'T'                              # Separator: T
            r'\d{2}:\d{2}:\d{2}'              # Time: hh:mm:ss
            r'(?:\.\d+)?'                     # Optional fractional seconds: .digits
            r'Z$'                             # Time zone: Z only
        )  
        
        self.formats_with_dots = [fmt for fmt in self.date_formats if '.' in fmt]
        self.formats_without_dots = [fmt for fmt in self.date_formats if '.' not in fmt]
        self.last_successful_format = None
        
        self.ET = self.EasternTime(self)         
        self.sub = self.SubstituteClass(self)          
        
        self.quick_parse = self.parse
        self.pro_parse = self.conversion        

    def _is_date_str(self, text):
        """Return whether the string contains a recognized date/time format directive."""
        date_format_string_pattern = re.compile(r"%[aAbBcdHImMpSUwWxXyYZ]")
        if not isinstance(text, str):
            return False
        return bool(date_format_string_pattern.search(text))

    def _is_datetimeType(self, obj, format='%Y-%m-%d', strf=False):
        """Identify date-like and datetime-like objects by their attributes.

        With strf=True, return a recognized object's strftime(format) result;
        otherwise return whether the attributes match.
        """
        date_attrs = {'year', 'month', 'day'}
        datetime_attrs = date_attrs.union({'hour', 'minute', 'second', 'microsecond'})

        if all(hasattr(obj, attr) for attr in datetime_attrs):
            if strf:
                return obj.strftime(format)
            return True
        elif all(hasattr(obj, attr) for attr in date_attrs) and not any(hasattr(obj, attr) for attr in datetime_attrs - date_attrs):
            if strf:
                return obj.strftime(format)
            return True
        return False

    def _from_pywintypes_datetime(self, obj):
        """Convert a pywintypes or datetime-like value to a standard datetime.

        Return other values unchanged.
        """
        if ((type(obj).__module__ == "pywintypes" and type(obj).__name__ == "datetime") or
            (hasattr(obj, "year") and hasattr(obj, "microsecond") and (hasattr(obj, "tzinfo") or hasattr(obj, "day")))):
            tzinfo = getattr(obj, 'tzinfo', None)
            return dt(
                obj.year,
                obj.month,
                obj.day,
                obj.hour,
                obj.minute,
                obj.second,
                obj.microsecond,
                tzinfo=tzinfo
            )
        return obj

    def _is_iso(self, text):
        """
        Check if the given text is in ISO 8601 format (UTC only).

        Parameters:
            text (str): The string to check.

        Returns:
            bool: True if the text matches the ISO 8601 format with a 'Z' timezone indicator; otherwise, False.
        """
        if not isinstance(text, str):
            return False

        return bool(re.compile(
            r'^'                  # Start of string
            r'\d{4}-\d{2}-\d{2}'  # Date: YYYY-MM-DD
            r'T'                  # Separator: T
            r'\d{2}:\d{2}:\d{2}'  # Time: hh:mm:ss
            r'(?:\.\d+)?'         # Optional fractional seconds: .digits
            r'Z$'                 # Time zone: Z only
        ).search(text))

    def _add_missing_seconds(self, date_str):
        """Add seconds to a time string only when its date prefix is parseable.

        Whitespace is normalized first. An invalid date prefix leaves the input
        unchanged; a recognized time without seconds receives :00.
        """
        cleaned = " ".join(date_str.split())
        
        pattern = r'^(?P<prefix>.*?)\s*(?P<time>(?P<hours>\d{1,2}):(?P<minutes>\d{2})(?::(?P<seconds>\d{2})(?:\.(?P<microseconds>\d+))?)?)\s*$'
        time_regex = re.compile(pattern)
        match = time_regex.search(cleaned)
        
        if not match:
            return date_str

        prefix = match.group("prefix")
        if prefix:
            try:
                parsed_date = self.parse(prefix)
            except Exception:
                return date_str
            if not isinstance(parsed_date, dt):
                return date_str

        hours = match.group("hours")
        minutes = match.group("minutes")
        seconds = match.group("seconds")
        microseconds = match.group("microseconds")
        
        if seconds is None:
            seconds = "00"
        
        new_time_str = f"{hours}:{minutes}:{seconds}"
        if microseconds:
            new_time_str += f".{microseconds}"
        
        if prefix:
            new_str = prefix.strip() + " " + new_time_str
        else:
            new_str = new_time_str
            
        return new_str


    def parse(self, date_input, from_format=None, to_format=None, to_unix_timestamp=False, include_timezone=False, timezone_offset='+00:00', keep_time=True):
        """
        Parse and convert dates from various formats.

        This method handles date parsing for multiple input types (e.g., string, list, numpy.ndarray, pandas.Series)
        and supports conversion to Unix timestamps, formatting, and ISO 8601 parsing.

        Parameters:
            date_input: A date input (string, datetime, list, numpy.ndarray, or pandas.Series).
            from_format (str, optional): Format string for parsing the input date.
            to_format (str, optional): Format string for output formatting.
            to_unix_timestamp (bool, optional): If True, returns an integer Unix timestamp.
            include_timezone (bool, optional): If True, includes a timezone offset for ISO dates.
            timezone_offset (str, optional): Timezone offset to use for ISO dates (default is '+00:00').
            keep_time (bool, optional): If False, returns only the date portion; otherwise, returns the full datetime.

        Returns:
            datetime.datetime, int, str, or collection: The parsed date in the specified format or as a Unix timestamp.

        Raises:
            ValueError: If the date format is not recognized.
        """
        def _parser(date_string):
            date_str = " ".join(date_string.split())            
            
            if self.last_successful_format:
                try:
                    parsed_date = dt.strptime(date_str, self.last_successful_format)
                    if to_unix_timestamp:
                        return int(parsed_date.timestamp())
                    if not from_format and self._is_date_str(to_format):
                        return parsed_date.strftime(to_format)
                    return parsed_date
                except ValueError:
                    pass
            for format_list in [self.formats_with_dots, self.formats_without_dots]:
                for date_format in format_list:
                    try:
                        parsed_date = dt.strptime(date_str, date_format)
                        self.last_successful_format = date_format
                        if to_unix_timestamp:
                            return int(parsed_date.timestamp())
                        if not from_format and self._is_date_str(to_format):
                            return parsed_date.strftime(to_format)
                        return parsed_date
                    except ValueError:
                        continue
            new_separators = ['/', '-']
            for sep in new_separators:
                for date_format in self.formats_with_dots:
                    new_format = date_format.replace('.', sep)
                    try:
                        parsed_date = dt.strptime(date_str, new_format)
                        self.last_successful_format = new_format
                        if to_unix_timestamp:
                            return int(parsed_date.timestamp())
                        if not from_format and self._is_date_str(to_format):
                            return parsed_date.strftime(to_format)
                        return parsed_date
                    except ValueError:
                        continue
            if from_format and to_format:
                try:
                    parsed_date = dt.strptime(date_str, from_format)
                    if to_unix_timestamp:
                        return int(parsed_date.timestamp())
                    formatted_date = parsed_date.strftime(to_format)
                    return formatted_date
                except ValueError:
                    raise ValueError("Date format not recognized and fallback failed.")
            elif from_format:
                try:
                    parsed_date = dt.strptime(date_str, from_format)
                    if to_unix_timestamp:
                        return int(parsed_date.timestamp())
                    return parsed_date
                except ValueError:
                    raise ValueError("Date format not recognized.")
            raise ValueError("Date format not recognized.")
        
        
        date_input = self._from_pywintypes_datetime(date_input)     
        if from_format is None:
            iso = try_iso(date_input)
            if iso is not None:
                # This compatibility parser may strip or replace the trailing Z offset.
                # Core Graph parsing uses parse_graph_datetime to retain timezone information.
                if isinstance(date_input, str) and date_input.endswith('Z'):
                    iso = iso.replace(tzinfo=offset_timezone(timezone_offset)) if include_timezone else iso.replace(tzinfo=None)
                if to_unix_timestamp: return int(iso.timestamp())
                if to_format: return iso.strftime(to_format)
                return iso if keep_time else iso.date()
        if not self._is_datetimeType(date_input) and isinstance(date_input, str):
            date_input = " ".join(date_input.split()) 
            date_input = self._add_missing_seconds(date_input)
        
        if self._is_datetimeType(date_input):
            if to_unix_timestamp:
                return int(date_input.timestamp())
            if to_format and self._is_date_str(to_format):
                return date_input.strftime(to_format)
            return date_input        
        
        if self._is_iso(date_input):
            if include_timezone:
                date_input = date_input.replace('Z', timezone_offset)
            else:
                date_input = date_input.replace('Z', '')
            formatted_date = dt.fromisoformat(date_input)
            if to_format:
                return formatted_date.strftime(to_format)
            return formatted_date if keep_time else formatted_date.date()        
                
        try:
            if isinstance(date_input, str):
                return _parser(date_input)
               
            elif isinstance(date_input, list) or isinstance(date_input, np.ndarray):
                return [self.parse(value, from_format, to_format, to_unix_timestamp, include_timezone, timezone_offset, keep_time) for value in date_input]
               
            elif sys.modules.get("pandas") is not None and isinstance(date_input, sys.modules["pandas"].Series):
                return date_input.map(lambda value: self.parse(value, from_format, to_format, to_unix_timestamp, include_timezone, timezone_offset, keep_time))
            else:
                raise ValueError("Unsupported basic date input")
            
                   
        except ValueError as e:
            raise ValueError("Date format not recognized; use pro_parse with m365[dates] for advanced formats") from e


    class EasternTime:
        def __init__(self, parent):
            self.parent = parent
            
        def isDST(self, date_value=None):
            """
            Determine if a given date is in Daylight Saving Time (DST) for Eastern Time.

            Parameters:
                date_value (datetime.datetime, optional): The date to check. If None, the current UTC time is used.

            Returns:
                bool: True if the date is within the DST period for Eastern Time; otherwise, False.
            """        	
            date_value = date_value if date_value else dt.utcnow()
            dst_start = dt(date_value.year, 3, 8)
            dst_end = dt(date_value.year, 11, 1) 
            while dst_start.weekday() != 6: 
                dst_start += td(days=1)
            while dst_end.weekday() != 6:
                dst_end += td(days=1)
            dst_start = dst_start.replace(hour=2)
            dst_end = dst_end.replace(hour=2)
            return dst_start <= date_value < dst_end 

        def now(self):
            """
            Get the current Eastern Time, accounting for Daylight Saving Time.

            Returns:
                datetime.datetime: The current time in Eastern Time (EDT or EST, as appropriate).
            """        	
            now_utc = dt.utcnow()
            year = now_utc.year
            dst_start = dt(year, 3, 8, 2) + td(days=(6 - dt(year, 3, 8, 2).weekday()))
            dst_end = dt(year, 11, 1, 2) + td(days=(6 - dt(year, 11, 1, 2).weekday()))
            if dst_start <= now_utc.replace(tzinfo=None) < dst_end:
                offset = td(hours=-4) # Eastern Daylight Time (UTC-4)
            else:
                offset = td(hours=-5) # Eastern Standard Time (UTC-5)
            now_est_edt = now_utc + offset
            return now_est_edt

        def nowOffset(self, datetime_datetime_obj=None):
            """
            Get the UTC offset in hours for Eastern Time.

            Parameters:
                datetime_datetime_obj (datetime.datetime, optional): A datetime object with timezone info.
                    If provided, its UTC offset is used; otherwise, the current time is evaluated.

            Returns:
                float: The UTC offset in hours (-4 for EDT or -5 for EST).
            """        	
            def has_timezone(dte):
                if isinstance(dte, dt):
                    return isinstance(dt.tzinfo, tmz)
                else:
                    return None
            if datetime_datetime_obj:
                if has_timezone(datetime_datetime_obj):
                    offset = datetime_datetime_obj.utcoffset()
                    offset_seconds = offset.total_seconds()
                    return offset_seconds/3600
            now_utc = dt.utcnow()
            year = now_utc.year
            dst_start = dt(year, 3, 8, 2) + td(days=(6 - dt(year, 3, 8, 2).weekday()))
            dst_end = dt(year, 11, 1, 2) + td(days=(6 - dt(year, 11, 1, 2).weekday()))
            if dst_start <= now_utc.replace(tzinfo=None) < dst_end:
                return -4 # Eastern Daylight Time (UTC-4)
            else:
                return -5 # Eastern Standard Time (UTC-5)
            return 0
    
        def __dir__(self):
            return ['isDST', 'now', 'nowOffset'] 
    
    
    class SubstituteClass:
        def __init__(self, parent):        
            """Initialize the time and date component helpers."""
            self.parent = parent            
            self.time = self.TimeClass(self.parent).time            
            self.date = self.DateClass().date               
          
        class TimeClass:
            def __init__(self, parent):
                self.parent = parent
        	
            def time(self, date_str, hours=None, minutes=None, seconds=None, microseconds=None):
                """
                1) Splits 'date_str' into (potential) 'prefix' (date-like substring) + 'time' (HH:MM, HH:MM:SS, or HH:MM:SS.microseconds).
                2) Uses '_dateplumb.pro_parse' to validate if the prefix or entire string is a parsable date.
                   - If no time portion is found, tries parsing the entire string as a date.
                3) If valid, reconstructs time components (hours, minutes, optional seconds, optional microseconds).
                   - If no original time was present, default missing components to "00", or build from user overrides.
                4) Returns the string with updated (or newly added) time if match+parse is successful; otherwise, returns 'date_str'.

                :param date_str: The original string containing a date-like prefix and optionally a time substring.
                :param hours: (str or int) Override hours, or None to use the matched hours (or "00" if no match).
                :param minutes: (str or int) Override minutes, or None to use the matched minutes (or "00" if no match).
                :param seconds: (str or int) Override seconds, or None to use the matched seconds (or none if no match).
                :param microseconds: (str or int) Override microseconds, or None to use the matched microseconds.
                :return: A modified string with updated or newly added time components if possible,
                         otherwise 'date_str' unchanged.
                """
                pattern = (
                    r'^(?P<prefix>.*?)\s*'                # capture anything (non-greedy) until whitespace
                    r'(?P<time>'                          # capture the time portion
                    r'(?P<hours>\d{1,2}):(?P<minutes>\d{2})'
                    r'(?:'
                    r':(?P<seconds>\d{2})'
                    r'(?:\.(?P<microseconds>\d+))?'
                    r')?'
                    r')\s*$'
                )
                time_regex = re.compile(pattern)

                cleaned_date_str = " ".join(date_str.split())

                match = time_regex.search(cleaned_date_str)
                
                if match:
                    prefix = match.group("prefix")
                    try:
                        parsed_date = self.parent.parse(prefix)
                    except Exception:
                        return date_str

                    if not isinstance(parsed_date, dt):
                        return date_str

                    parts = match.groupdict()
                    found_hours = parts.get("hours")
                    found_minutes = parts.get("minutes")
                    found_seconds = parts.get("seconds")
                    found_microseconds = parts.get("microseconds")

                    # Keep the matched components unless an override is supplied.
                    final_hours = str(hours) if hours is not None else found_hours
                    final_minutes = str(minutes) if minutes is not None else found_minutes
                    final_seconds = str(seconds) if seconds is not None else found_seconds

                    final_microseconds = str(microseconds) if microseconds is not None else found_microseconds

                    # Ensure microseconds are always six digits if present
                    if final_microseconds is not None:
                        final_microseconds = final_microseconds.ljust(6, '0')

                    new_time_str = f"{final_hours}:{final_minutes}"
                    if final_seconds is not None:
                        new_time_str += f":{final_seconds}"
                        if final_microseconds is not None:
                            new_time_str += f".{final_microseconds}"

                    time_start, time_end = match.span("time")
                    new_str = cleaned_date_str[:time_start] + new_time_str + cleaned_date_str[time_end:]
                    return new_str

                else:
                    try:
                        parsed_date = self.parent.parse(cleaned_date_str)                    	
                    except Exception:
                        return date_str

                    if not isinstance(parsed_date, dt):
                        return date_str

                    final_hours = str(hours) if hours is not None else "00"
                    final_minutes = str(minutes) if minutes is not None else "00"
                    final_seconds = str(seconds) if seconds is not None else "00"
                    final_microseconds = None
                    if microseconds is not None:
                        final_microseconds = str(microseconds).ljust(6, '0')

                    new_time_str = f"{final_hours}:{final_minutes}:{final_seconds}"
                    if final_microseconds:
                        new_time_str += f".{final_microseconds}"

                    return f"{cleaned_date_str} {new_time_str}"

        class DateClass:
            def date(self, date_strings, year=None, month=None, day=None):
                """
                Replace specific components in a date string or a collection of date strings with new values.
                 
                This function parses a date string to identify existing components and replaces them with
                new values provided as arguments. It reconstructs the date string with the new values.
                 
                Parameters:
                    date_strings (str, list, np.ndarray, pd.Series): The date string or collection to modify.
                    year (str or int or None): The new year value to replace the existing year.
                    month (str or int or None): The new month value to replace the existing month.
                    day (str or int or None): The new day value to replace the existing day.
                 
                Returns:
                    str, list, np.ndarray, pd.Series: The modified date string or collection with the new values.
                """                
                return _advanced.replace_date(date_strings, year=year, month=month, day=day)

        def __dir__(self):
            return ['time', 'date'] 
           
           
    def now(self, utc=False, as_unix=False, format=None):
        """
        Get the current date and time.

        Parameters:
            utc (bool, optional): If True, returns the current UTC time instead of local time. Defaults to False.
            as_unix (bool, optional): If True, returns the Unix timestamp instead of a datetime object. Defaults to False.
            format (str, optional): If provided, returns the formatted date string according to this format.

        Returns:
            datetime.datetime | int | str: 
                - If `as_unix` is True, returns an integer Unix timestamp.
                - If `format` is provided, returns a formatted string representation of the date.
                - Otherwise, returns a `datetime.datetime` object.
        """
        current = dt.utcnow() if utc else dt.now()
        if as_unix:
            return self.unix_timestamp(current)
        return current.strftime(format) if format else current

    def nowCT(self, as_unix=False, format=None):
        """
        Retrieve the current date and time in Central Time (CT).

        Parameters:
            as_unix (bool, optional): If True, return the Unix timestamp. Defaults to False.
            format (str, optional): Format string for output. Defaults to None.

        Returns:
            datetime.datetime, int, or str: The current CT time as a datetime object, Unix timestamp, or formatted string.
        """
        now_utc = dt.utcnow()
        if self.EasternTime.isDST(now_utc):           
            current = now_utc - td(hours=5)  # UTC-5 for DST
        else:
            current = now_utc - td(hours=6)  # UTC-6 for Standard Time
        if as_unix:
            return self.unix_timestamp(current)
        return current.strftime(format) if format else current

    def make_timezone_aware(self, date_value, offset_hours):
        """
        Make a naive datetime object timezone-aware.

        Parameters:
            date_value (datetime.datetime): A naive datetime object.
            offset_hours (int or float): UTC offset in hours to apply.

        Returns:
            datetime.datetime: The timezone-aware datetime object.
        """
        if date_value.tzinfo is not None:
            return date_value
        tz = tmz(td(hours=offset_hours))
        return date_value.replace(tzinfo=tz)

    def unix_timestamp(self, date_value, format=None, utc=False, reset_time=False, to_unix=True):
        """
        Convert between Unix timestamps and datetime objects.

        Parameters:
            date_value (datetime.datetime or int):
                - If converting to Unix (to_unix=True), it must be a datetime object.
                - If converting from Unix (to_unix=False), it must be an integer timestamp.
            format (str, optional): Format string for output when converting from Unix timestamp. Defaults to None.
            utc (bool, optional): If True, treat the datetime as UTC. Defaults to False.
            reset_time (bool, optional): If True, reset the time to 00:00:00. Defaults to False.
            to_unix (bool, optional): Direction of conversion. True for datetime to Unix, False for Unix to datetime.

        Returns:
            int, datetime.datetime, or str: The Unix timestamp, datetime object, or formatted date string.

        Raises:
            ValueError: If the type of `date_value` is not as expected.
        """
        if not isinstance(date_value, (int, dt)):
            try:
                date_value = self.parse(date_value)
            except:
                raise ValueError("Expected datetime object or integer for Unix timestamp conversion.")          
        
        if to_unix:
            if not isinstance(date_value, dt):
                raise ValueError("Expected datetime object for Unix timestamp conversion.")
            datetime_obj = date_value
            if utc:
                datetime_obj = datetime_obj.replace(tzinfo=tmz.utc)
            if reset_time:
                datetime_obj = datetime_obj.replace(hour=0, minute=0, second=0, microsecond=0)
            return int(datetime_obj.timestamp())
        else:
            if not isinstance(date_value, int):
                raise ValueError("Expected integer Unix timestamp for conversion to datetime.")
            if utc:
                datetime_obj = dt.fromtimestamp(date_value, tmz.utc)
            else:
                datetime_obj = dt.fromtimestamp(date_value)
            return datetime_obj.strftime(format) if format else datetime_obj

    def subtract_months(self, date_str, months):
        """
        Subtract a specified number of months from a given date.

        Parameters:
            date_str (str): The input date as a string in the format '%Y-%m-%d'.
            months (int): The number of months to subtract.

        Returns:
            str: The resulting date string formatted as '%Y-%m-01'.
        """  	
        from calendar import monthrange    
        date = dt.strptime(date_str, '%Y-%m-%d')
        y, m = divmod(date.year * 12 + (date.month - 1) - months, 12)
        y, m = y, m + 1
        day = min(date.day, monthrange(y, m)[1])
        return dt(y, m, day).strftime('%Y-%m-01')        

    def days_in_year(self, year):
        """
        Calculate the number of days in a given year.

        Parameters:
            year (int): The year to evaluate.

        Returns:
            int: 366 if the year is a leap year; otherwise, 365.
        """
        if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
            return 366
        else:
            return 365

    def weeks_in_year(self, date=None, mode='passed'):
        """
        Calculate the number of weeks passed or remaining in the year from a given date.

        Parameters:
            date (str or datetime.date, optional): The reference date (in '%Y-%m-%d' format or as a date object). Defaults to today.
            mode (str, optional): Calculation mode - 'passed' for weeks passed, 'left' for weeks remaining. Defaults to 'passed'.

        Returns:
            int: The number of weeks based on the specified mode.

        Raises:
            ValueError: If an invalid mode is provided.
        """
        if date is None:
            date = d.today()
        elif isinstance(date, str):
            date = self.parse(date).date()

        start_of_year = d(date.year, 1, 1)
        end_of_year = d(date.year, 12, 31)

        if mode == 'passed':
            days_passed = (date - start_of_year).days
            return days_passed // 7
        elif mode == 'left':
            days_left = (end_of_year - date).days
            return days_left // 7
        else:
            raise ValueError("Invalid mode. Please choose 'passed' or 'left'.")

    def conversion(self, dates, to_format=None, delta=0, dict_keys=None, dict_inplace=False):
        """
        This function serves as a versatile converter for date and datetime inputs. It supports converting single or 
        multiple date strings or datetime objects into a specified format or datetime objects, with the option to modify
        the date by a given delta of days. Additionally, it handles dictionaries containing date information by applying
        conversions recursively to specified keys.
         
        This function is particularly useful in data preprocessing where dates might come in various formats and need
        standardization or adjustment based on a temporal delta for further analysis or storage.
         
        Parameters:
            dates (str, list, np.ndarray, pd.Series, datetime.datetime, dict): The input date(s) which can be a single date string,
                a datetime object, a collection (list, array, series) of date strings or datetime objects, or a dictionary containing
                date strings or datetime objects nested under specified keys.
            to_format (str, optional): The desired output format of the date(s) as a string according to datetime.strftime conventions.
                If None, the function will return datetime objects instead of formatted strings.
            delta (int, default=0): An integer representing the number of days to add or subtract from the input date(s). Positive
                values move the date forward, while negative values move it backwards.
            dict_keys (list, optional): When the 'dates' parameter is a dictionary, this list specifies which keys contain the
                date information to be converted. This parameter is mandatory if 'dates' is a dictionary.
            dict_inplace (bool, default=False): Determines whether the dictionary is modified in place. If True, the dictionary is
                modified directly and the function returns None. If False, the function modifies a copy of the dictionary and returns it.
         
        Returns:
            The function returns the converted date(s) either as formatted strings (if 'to_format' is specified) or as datetime
                objects. If the input is a dictionary and 'dict_inplace' is False, it returns a new dictionary with the dates converted.
                If 'dict_inplace' is True, the input dictionary is modified directly and the function returns None.
         
        Raises:
            ValueError: If 'dates' is a dictionary and 'dict_keys' is not provided, or if any input date string format is unrecognized
                or incorrect, making it impossible to parse the date.
        """        
        return _advanced.conversion(dates, to_format, delta, dict_keys, dict_inplace)

    def detect(self, date_strings):
        """
        Detect and adjust the date format based on the components of a single date string or a collection of date strings.
        This function analyzes each date string to identify its format and makes adjustments to handle leading zeros in the date components.
        It ensures consistent date formatting across different platforms by replacing zero-padded specifiers with their non-zero-padded
        counterparts where applicable.
         
        Parameters:
            date_strings (str, list, np.ndarray, pd.Series): The date string or collection of date strings to analyze.
         
        Returns:
            str, list, np.ndarray, pd.Series: Depending on the input type, returns either a single format or a collection of formats with detected and possibly adjusted date format strings.
         
        Raises:
            ValueError: If no matching format is found for any of the given date strings.
        """      
        return _advanced.detect(date_strings)

    def fragment(self, date_strings, component, ret_format=False):
        """
        Extract specific date components from a single date string or a collection of date strings using precompiled regex patterns.
         
        This function parses a date string to extract specified components, ensuring accurate extraction
        of year, month, day, hour, minute, and second components, which are critical for consistent date
        formatting across different platforms.
         
        Parameters:
            date_strings (str): The date string to parse.
            component (str): The component to extract. Valid components include:
                'year', 'month', 'day', 'hour24', 'hour12', 'minutes', 'seconds'.
         
        Returns:
            str or None: The extracted component as a string, or None if no match is found.
         
        Components:
            - 'year': Year component of the date.
            - 'month': Month component of the date.
            - 'day': Day component of the date.
            - 'weekday': Weekday component of the date.
            - 'hour24': Hour component in 24-hour format.
            - 'hour12': Hour component in 12-hour format.
            - 'minute': Minute component of the time.
            - 'second': Second component of the time.
            - 'microsecond': Microsecond component of the time.
        """        
        return _advanced.fragment(date_strings, component, ret_format)

    def sequence(self, start_date, end_date, to_format=None):
        """
        Generate a sequence of formatted dates between two dates.
         
        Parameters:
            start_date (any): The start date of the sequence, which can be a datetime object or convertible to one.
            end_date (any): The end date of the sequence, which can be a datetime object or convertible to one.
            to_format (str, optional): The format string to use for formatting the dates.
                Defaults to None.
         
        Returns:
            list of str: List of formatted date strings from start to end date inclusive.
        """  
        return _advanced.sequence(start_date, end_date, to_format)

    def contains_time(self, obj):
        """
        Determines whether the provided object has non-zero time information.

        Accepts datetime, time, or date objects.
        
        - For datetime and time objects, returns True if at least one of the time components
          (hour, minute, second, microsecond) is non-zero.
        - For date objects (which lack time data), returns False.
        
        :param obj: A datetime, date, or time object.
        :return: True if the object has non-zero time information; otherwise, False.
        :raises TypeError: If 'obj' is not a recognized date/time type.
        """
        if isinstance(obj, (dt, t)):
            return (obj.hour, obj.minute, obj.second, obj.microsecond) != (0, 0, 0, 0)
        return False

    def __dir__(self):
        return ['parse', 'now', 'unix_timestamp', 'nowCT', 'subtract_months', 'days_in_year', 'weeks_in_year', 'ET', 'make_timezone_aware', 'conversion', 'detect', 'fragment', 'sequence', 'sub', 'quick_parse', 'pro_parse', 'contains_time']
       
dtparse = dt_parse()

def __dir__():
    return ['dtparse']

__all__ = ['dtparse']
