use core::fmt;
use std::ops;

pub const EPOCH: i32 = 1900;

pub const MS_PER_SEC: i64 = 1000;
pub const MS_PER_MIN: i64 = MS_PER_SEC * 60;
pub const MS_PER_HOUR: i64 = MS_PER_MIN * 60;
pub const MS_PER_DAY: i64 = MS_PER_HOUR * 24;

pub const MONTH_LENGTHS: [u8; 12] = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum ParseError {
    InvalidFormat,
    InvalidDate,
    InvalidTime,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct DurationParts {
    pub days: i64,
    pub hours: u8,
    pub minutes: u8,
    pub seconds: u8,
    pub milliseconds: u16,
}

impl DurationParts {
    pub fn new(days: i64, hours: u8, minutes: u8, seconds: u8, milliseconds: u16) -> DurationParts {
        DurationParts {
            days,
            hours,
            minutes,
            seconds,
            milliseconds,
        }
    }
}

impl From<&Duration> for DurationParts {
    fn from(duration: &Duration) -> DurationParts {
        let mut remaining = duration.0;

        let d = remaining / MS_PER_DAY;
        remaining -= d * MS_PER_DAY;

        let h = (remaining / MS_PER_HOUR) as u8;
        remaining -= h as i64 * MS_PER_HOUR;

        let m = (remaining / MS_PER_MIN) as u8;
        remaining -= m as i64 * MS_PER_MIN;

        let s = (remaining / MS_PER_SEC) as u8;
        remaining -= s as i64 * MS_PER_SEC;

        let ms = remaining as u16;
        DurationParts::new(d, h, m, s, ms)
    }
}

impl fmt::Display for DurationParts {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let sign = if self.days < 0 { "-" } else { "" };
        write!(
            f,
            "{sign}{}:{:02}:{:02}:{:02}.{:03}",
            self.days, self.hours, self.minutes, self.seconds, self.milliseconds
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct DateTimeParts {
    pub year: i32,
    pub month: u8,
    pub day: u8,
    pub hour: u8,
    pub minute: u8,
    pub second: u8,
    pub millisecond: u16,
}

impl DateTimeParts {
    pub fn new(
        year: i32,
        month: u8,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
        millisecond: u16,
    ) -> DateTimeParts {
        DateTimeParts {
            year,
            month,
            day,
            hour,
            minute,
            second,
            millisecond,
        }
    }
}

impl From<&DateTime> for DateTimeParts {
    fn from(excel_date: &DateTime) -> DateTimeParts {
        let (year, month, day) = excel_date.date_parts();
        let (hour, minute, second, millisecond) = excel_date.time_parts();
        DateTimeParts::new(year, month, day, hour, minute, second, millisecond)
    }
}

impl TryFrom<&str> for DateTimeParts {
    type Error = ParseError;

    /// Parses a datetime string in the format `(-)YYYY-MM-DDTHH:MM:SS.SSS` into a `DateTimeParts` struct.
    ///
    /// # Errors
    ///
    /// `ParseError::InvalidFormat` if the input string is empty or does not contain a valid datetime format.
    ///
    /// `ParseError::InvalidDate` if any of the date parts fall outside range.
    ///
    /// `ParseError::InvalidTime` if any of the time parts fall outside range.
    ///
    /// # Examples
    ///
    /// ```
    /// use excel_xml::datetime::DateTimeParts;
    ///
    /// let a = DateTimeParts::try_from("2020-01-01T00:00:00.000").unwrap();
    /// assert_eq!(a.year, 2020);
    /// assert_eq!(a.hour, 0);
    ///
    /// let b = DateTimeParts::try_from("-2020-01-01T01:00:00.010").unwrap();
    /// assert_eq!(b.year, -2020);
    /// assert_eq!(b.hour, 1);
    /// assert_eq!(b.millisecond, 10);
    /// ```
    fn try_from(value: &str) -> Result<Self, Self::Error> {
        let bytes = value.as_bytes();
        if bytes.is_empty() {
            return Err(ParseError::InvalidFormat);
        }

        let mut idx = 0;

        if bytes[idx] == b'-' {
            idx += 1;
        }
        let year_start = idx;
        while idx < bytes.len() && bytes[idx].is_ascii_digit() {
            idx += 1;
        }
        if idx == year_start {
            return Err(ParseError::InvalidFormat);
        }

        let slice = |a: usize, b: usize| value.get(a..b).ok_or(ParseError::InvalidFormat);

        let year = match slice(0, idx)?.parse::<i32>() {
            Ok(y) if y != 0 => y,
            Ok(_) => return Err(ParseError::InvalidDate),
            Err(_) => return Err(ParseError::InvalidFormat),
        };

        if bytes.get(idx) != Some(&b'-') {
            return Err(ParseError::InvalidFormat);
        }
        idx += 1;
        let month = match slice(idx, idx + 2)?.parse::<u8>() {
            Ok(m) if m >= 1 && m <= 12 => m,
            Ok(_) => return Err(ParseError::InvalidDate),
            Err(_) => return Err(ParseError::InvalidFormat),
        };

        if bytes.get(idx + 2) != Some(&b'-') {
            return Err(ParseError::InvalidFormat);
        }
        idx += 3;
        let day = match slice(idx, idx + 2)?.parse::<u8>() {
            Ok(d) if d >= 1 && d <= 31 => d,
            Ok(_) => return Err(ParseError::InvalidDate),
            Err(_) => return Err(ParseError::InvalidFormat),
        };
        if day > get_days_in_month(month, year) {
            return Err(ParseError::InvalidDate);
        }

        if bytes.get(idx + 2) != Some(&b'T') {
            return Err(ParseError::InvalidFormat);
        }
        idx += 3;
        let hour = match slice(idx, idx + 2)?.parse::<u8>() {
            Ok(h) if h <= 23 => h,
            Ok(_) => return Err(ParseError::InvalidTime),
            Err(_) => return Err(ParseError::InvalidFormat),
        };

        if bytes.get(idx + 2) != Some(&b':') {
            return Err(ParseError::InvalidFormat);
        }
        idx += 3;
        let minute = match slice(idx, idx + 2)?.parse::<u8>() {
            Ok(m) if m <= 59 => m,
            Ok(_) => return Err(ParseError::InvalidTime),
            Err(_) => return Err(ParseError::InvalidFormat),
        };

        if bytes.get(idx + 2) != Some(&b':') {
            return Err(ParseError::InvalidFormat);
        }
        idx += 3;
        let second = match slice(idx, idx + 2)?.parse::<u8>() {
            Ok(s) if s <= 59 => s,
            Ok(_) => return Err(ParseError::InvalidTime),
            Err(_) => return Err(ParseError::InvalidFormat),
        };

        if bytes.get(idx + 2) != Some(&b'.') {
            return Err(ParseError::InvalidFormat);
        }
        idx += 3;
        let millisecond = match slice(idx, idx + 3)?.parse::<u16>() {
            Ok(ms) if ms <= 999 => ms,
            Ok(_) => return Err(ParseError::InvalidTime),
            Err(_) => return Err(ParseError::InvalidFormat),
        };

        if idx + 3 != bytes.len() {
            return Err(ParseError::InvalidFormat);
        }
        Ok(DateTimeParts::new(
            year,
            month,
            day,
            hour,
            minute,
            second,
            millisecond,
        ))
    }
}

impl fmt::Display for DateTimeParts {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}",
            self.year, self.month, self.day, self.hour, self.minute, self.second, self.millisecond
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Duration(i64);

impl Duration {
    pub fn new(days: i64, hours: i64, minutes: i64, seconds: i64, milliseconds: i64) -> Duration {
        Duration(
            (days * MS_PER_DAY)
                + (hours * MS_PER_HOUR)
                + (minutes * MS_PER_MIN)
                + (seconds * MS_PER_SEC)
                + milliseconds,
        )
    }

    pub fn from_millis(m: i64) -> Duration {
        Duration::new(0, 0, 0, 0, m)
    }

    pub fn from_seconds(s: i64) -> Duration {
        Duration::new(0, 0, 0, s, 0)
    }

    pub fn from_minutes(m: i64) -> Duration {
        Duration::new(0, 0, m, 0, 0)
    }

    pub fn from_hours(h: i64) -> Duration {
        Duration::new(0, h, 0, 0, 0)
    }

    pub fn from_days(d: i64) -> Duration {
        Duration::new(d, 0, 0, 0, 0)
    }

    pub fn saturating_add(&self, other: Duration) -> Duration {
        Duration(self.0.saturating_add(other.0))
    }

    pub fn saturating_sub(&self, other: Duration) -> Duration {
        Duration(self.0.saturating_sub(other.0))
    }

    pub fn as_millis(&self) -> i64 {
        self.0
    }

    pub fn as_seconds(&self) -> f64 {
        self.0 as f64 / MS_PER_SEC as f64
    }

    pub fn as_minutes(&self) -> f64 {
        self.0 as f64 / MS_PER_MIN as f64
    }

    pub fn as_hours(&self) -> f64 {
        self.0 as f64 / MS_PER_HOUR as f64
    }

    pub fn as_days(&self) -> f64 {
        self.0 as f64 / MS_PER_DAY as f64
    }

    pub fn parts(&self) -> DurationParts {
        DurationParts::from(self)
    }
}

impl From<DurationParts> for Duration {
    fn from(parts: DurationParts) -> Duration {
        Duration::new(
            parts.days,
            parts.hours as i64,
            parts.minutes as i64,
            parts.seconds as i64,
            parts.milliseconds as i64,
        )
    }
}

impl fmt::Display for Duration {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.parts())
    }
}

impl ops::Add for Duration {
    type Output = Duration;

    fn add(self, rhs: Duration) -> Duration {
        Duration(self.0 + rhs.0)
    }
}

impl ops::Sub for Duration {
    type Output = Duration;

    fn sub(self, rhs: Duration) -> Duration {
        Duration(self.0 - rhs.0)
    }
}

impl ops::AddAssign for Duration {
    fn add_assign(&mut self, rhs: Duration) {
        *self = Duration(self.0 + rhs.0);
    }
}

impl ops::SubAssign for Duration {
    fn sub_assign(&mut self, rhs: Duration) {
        *self = Duration(self.0 - rhs.0);
    }
}

impl ops::Neg for Duration {
    type Output = Duration;

    fn neg(self) -> Duration {
        Duration(-self.0)
    }
}

impl ops::Mul<i64> for Duration {
    type Output = Duration;

    fn mul(self, rhs: i64) -> Duration {
        Duration(self.0 * rhs)
    }
}

impl ops::Div<i64> for Duration {
    type Output = Duration;

    fn div(self, rhs: i64) -> Duration {
        Duration(self.0 / rhs)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct DateTime(Duration);

impl DateTime {
    pub fn new(
        year: i32,
        month: u8,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
        millisecond: u16,
    ) -> DateTime {
        let days = calculate_year_offset(year)
            + calculate_month_offset(month, year)
            + (day - 1) as i64 * MS_PER_DAY;

        DateTime(Duration::new(
            days,
            hour as i64,
            minute as i64,
            second as i64,
            millisecond as i64,
        ))
    }

    pub fn from_date(year: i32, month: u8, day: u8) -> DateTime {
        DateTime::new(year, month, day, 0, 0, 0, 0)
    }

    pub fn duration_since_epoch(&self) -> Duration {
        self.0
    }

    pub fn date_parts(&self) -> (i32, u8, u8) {
        let (y, mut remaining) = self.get_year_and_remaining_days();

        let mut m = 1;
        loop {
            let days_in_month = get_days_in_month(m, y) as i64;
            if remaining < days_in_month {
                break;
            } else {
                remaining -= days_in_month;
                m += 1;
            }
        }

        let d = (remaining + 1) as u8;
        (y, m, d)
    }

    pub fn time_parts(&self) -> (u8, u8, u8, u16) {
        let mut remaining = self.0.0;

        let h = (remaining / MS_PER_HOUR) as u8;
        remaining -= h as i64 * MS_PER_HOUR;

        let m = (remaining / MS_PER_MIN) as u8;
        remaining -= m as i64 * MS_PER_MIN;

        let s = (remaining / MS_PER_SEC) as u8;
        remaining -= s as i64 * MS_PER_SEC;

        let ms = remaining as u16;
        (h, m, s, ms)
    }

    pub fn parts(&self) -> DateTimeParts {
        DateTimeParts::from(self)
    }

    pub fn is_leap_year(&self) -> bool {
        let (y, _) = self.get_year_and_remaining_days();
        is_leap_year(y)
    }

    fn get_year_and_remaining_days(&self) -> (i32, i64) {
        let mut remaining = self.0.as_days() as i64;

        let mut y = EPOCH;
        loop {
            let days_in_year = get_days_in_year(y) as i64;
            if remaining < days_in_year {
                break;
            } else {
                remaining -= days_in_year;
                y += 1;
            }
        }

        (y, remaining)
    }
}

impl From<DateTimeParts> for DateTime {
    fn from(parts: DateTimeParts) -> DateTime {
        DateTime::new(
            parts.year,
            parts.month,
            parts.day,
            parts.hour,
            parts.minute,
            parts.second,
            parts.millisecond,
        )
    }
}

impl TryFrom<&str> for DateTime {
    type Error = ParseError;

    /// Parses a string into a DateTime.
    ///
    /// # Errors
    ///
    /// `ParseError::InvalidFormat` if the input string is empty or does not contain a valid datetime format.
    ///
    /// `ParseError::InvalidDate` if any of the date parts fall outside range.
    ///
    /// `ParseError::InvalidTime` if any of the time parts fall outside range.
    ///
    /// # Examples
    ///
    /// ```
    /// use excel_xml::datetime::DateTime;
    ///
    /// let a = DateTime::try_from("2020-01-01T00:00:00.000").unwrap();
    /// assert_eq!(a.date_parts().year, 2020);
    ///
    /// let b = DateTime::try_from("-2020-01-01T01:00:00.000").unwrap();
    /// assert_eq!(b.date_parts().year, -2020);
    /// assert_eq!(b.date_parts().hour, 1);
    /// ```
    fn try_from(s: &str) -> Result<DateTime, Self::Error> {
        let parts = DateTimeParts::try_from(s)?;
        Ok(parts.into())
    }
}

impl fmt::Display for DateTime {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.parts())
    }
}

impl ops::Sub for DateTime {
    type Output = Duration;

    /// Subtracts two DateTime objects and returns a Duration.
    /// The result will be negative if `self` is earlier than `rhs`.
    ///
    /// # Examples
    ///
    /// ```
    /// use excel_xml::datetime::DateTime;
    ///
    /// let a = DateTime::from_date(2020, 3, 1);
    /// let b = DateTime::from_date(2020, 1, 1);
    /// let c = a - b;
    /// assert_eq!(c.as_days(), 60.0);
    ///
    /// let d = b - a;
    /// assert_eq!(d.as_days(), -60.0);
    /// ```
    fn sub(self, rhs: DateTime) -> Duration {
        self.0 - rhs.0
    }
}

impl ops::Add<Duration> for DateTime {
    type Output = DateTime;

    /// Adds a Duration to a DateTime and returns a new DateTime.
    /// Negative Durations can be used to subtract from a DateTime.
    ///
    /// # Examples
    ///
    /// ```
    /// use excel_xml::datetime::{DateTime, Duration};
    ///
    /// let a = DateTime::from_date(2020, 1, 1);
    /// let b = Duration::from_days(60);
    /// let c = a + b;
    /// assert_eq!(c, DateTime::from_date(2020, 3, 1));
    ///
    /// let d = Duration::from_days(-31);
    /// let e = a + d;
    /// assert_eq!(e, DateTime::from_date(2019, 12, 1));
    /// ```
    fn add(self, rhs: Duration) -> DateTime {
        DateTime(self.0 + rhs)
    }
}

impl ops::Sub<Duration> for DateTime {
    type Output = DateTime;

    /// Subtracts a Duration from a DateTime and returns a new DateTime.
    /// Negative Durations can be used to add to a DateTime.
    ///
    /// # Examples
    ///
    /// ```
    /// use excel_xml::datetime::{DateTime, Duration};
    ///
    /// let a = DateTime::from_date(2020, 3, 1);
    /// let b = Duration::from_days(60);
    /// let c = a - b;
    /// assert_eq!(c, DateTime::from_date(2020, 1, 1));
    ///
    /// let d = Duration::from_days(-31);
    /// let e = a - d;
    /// assert_eq!(e, DateTime::from_date(2020, 4, 1));
    /// ```
    fn sub(self, rhs: Duration) -> DateTime {
        DateTime(self.0 - rhs)
    }
}

impl ops::AddAssign<Duration> for DateTime {
    /// Adds a Duration to a DateTime and assigns the result to the DateTime.
    ///
    /// # Examples
    /// ```
    /// use excel_xml::datetime::{DateTime, Duration};
    ///
    /// let mut a = DateTime::from_date(2020, 1, 1);
    /// let b = Duration::from_days(60);
    /// a += b;
    /// assert_eq!(a, DateTime::from_date(2020, 3, 1));
    /// ```
    fn add_assign(&mut self, rhs: Duration) {
        *self = DateTime(self.0 + rhs);
    }
}

impl ops::SubAssign<Duration> for DateTime {
    /// Subtracts a Duration from a DateTime and assigns the result to the DateTime.
    ///
    /// # Examples
    /// ```
    /// use excel_xml::datetime::{DateTime, Duration};
    ///
    /// let mut a = DateTime::from_date(2020, 3, 1);
    /// let b = Duration::from_days(60);
    /// a -= b;
    /// assert_eq!(a, DateTime::from_date(2020, 1, 1));
    /// ```
    fn sub_assign(&mut self, rhs: Duration) {
        *self = DateTime(self.0 - rhs);
    }
}

pub fn is_leap_year(year: i32) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

/// Returns the number of days in a year
/// (366 for leap years, 365 otherwise)
pub fn get_days_in_year(year: i32) -> u16 {
    match is_leap_year(year) {
        true => 366,
        false => 365,
    }
}

/// Returns the number of days in a month
/// (29 for February in a leap year, 28 otherwise)
pub fn get_days_in_month(month: u8, year: i32) -> u8 {
    if month == 2 && is_leap_year(year) {
        29
    } else {
        MONTH_LENGTHS[month as usize - 1]
    }
}

fn calculate_year_offset(year: i32) -> i64 {
    match year < EPOCH {
        true => (year..EPOCH).fold(0, |acc, i| acc - get_days_in_year(i) as i64),
        false => (EPOCH..year).fold(0, |acc, i| acc + get_days_in_year(i) as i64),
    }
}

fn calculate_month_offset(month: u8, year: i32) -> i64 {
    (1..month).fold(0, |acc, i| acc + get_days_in_month(i, year) as i64)
}
