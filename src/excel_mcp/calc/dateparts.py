"""Excel's 1900 calendar, with its day 0 and its non-existent 29 February 1900."""

import datetime as dt

LAST_SERIAL = 2_958_465  # 31 December 9999


def serial_day(serial: float) -> int:
    """The day a serial number falls on once Excel has rounded it to the second."""
    return round(serial * 86400) // 86400


def month_day_year(serial: int) -> tuple[int, int, int]:
    """(year, month, day) of a serial number from 0 to LAST_SERIAL, as Excel shows it."""
    if serial == 0:
        return 1900, 1, 0
    if serial == 60:
        return 1900, 2, 29
    start = dt.date(1899, 12, 31) if serial < 60 else dt.date(1899, 12, 30)
    date = start + dt.timedelta(days=serial)
    return date.year, date.month, date.day


def day_of_week(serial: int) -> int:
    """0 for Sunday: serial 1, which Excel calls 1 January 1900, is a Sunday."""
    return (serial - 1) % 7
