// Time calculation utilities

/**
 * Parse a time string (HH:MM or HH:MM:SS) to minutes from midnight
 */
export function parseTimeToMinutes(time: string | null | undefined): number | null {
  if (!time || time === '' || time === '__' || time.includes('__')) return null;
  
  // Clean the time string
  const cleanTime = time.replace(/[^\d:]/g, '').trim();
  if (!cleanTime) return null;
  
  const parts = cleanTime.split(':');
  if (parts.length < 2) return null;
  
  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);
  
  if (isNaN(hours) || isNaN(minutes)) return null;
  
  return hours * 60 + minutes;
}

/**
 * Convert minutes from midnight to HH:MM format
 */
export function minutesToTimeString(minutes: number): string {
  if (minutes < 0) return '0:00';
  const hours = Math.floor(minutes / 60);
  const mins = Math.round(minutes % 60);
  return `${hours}:${mins.toString().padStart(2, '0')}`;
}

/**
 * Get the day of week (0 = Sunday, 1 = Monday, etc.)
 */
export function getDayOfWeek(date: Date): number {
  return date.getDay();
}

/**
 * Get day name abbreviation
 */
export function getDayName(date: Date): string {
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return days[date.getDay()];
}

/**
 * Check if a date is Friday (5)
 */
export function isFriday(date: Date): boolean {
  return date.getDay() === 5;
}

/**
 * Check if a date is Monday (1) or Thursday (4)
 */
export function isMondayOrThursday(date: Date): boolean {
  const day = date.getDay();
  return day === 1 || day === 4;
}

/**
 * Check if a date is a weekend (Saturday = 6, Sunday = 0)
 */
export function isWeekend(date: Date): boolean {
  const day = date.getDay();
  return day === 0 || day === 6;
}

/**
 * Calculate break overlap in minutes
 * Returns the number of minutes that overlap between the work period and break period
 */
export function calculateBreakOverlap(
  clockInMinutes: number,
  clockOutMinutes: number,
  breakStartMinutes: number,
  breakEndMinutes: number
): number {
  // No overlap if work doesn't include the break period
  if (clockOutMinutes <= breakStartMinutes || clockInMinutes >= breakEndMinutes) {
    return 0;
  }
  
  // Calculate the overlap
  const overlapStart = Math.max(clockInMinutes, breakStartMinutes);
  const overlapEnd = Math.min(clockOutMinutes, breakEndMinutes);
  
  return Math.max(0, overlapEnd - overlapStart);
}

/**
 * Parse date from various formats (DD/MM/YYYY, MM/DD/YYYY, etc.)
 */
export function parseDate(dateStr: string): Date | null {
  if (!dateStr) return null;
  
  // Try DD/MM/YYYY format first
  const ddmmyyyy = dateStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (ddmmyyyy) {
    const day = parseInt(ddmmyyyy[1], 10);
    const month = parseInt(ddmmyyyy[2], 10) - 1;
    const year = parseInt(ddmmyyyy[3], 10);
    return new Date(year, month, day);
  }
  
  // Try "Oct 1, 2025" format
  const monthDayYear = dateStr.match(/(\w+)\s+(\d{1,2}),?\s+(\d{4})/);
  if (monthDayYear) {
    const monthNames: Record<string, number> = {
      'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
      'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
    };
    const month = monthNames[monthDayYear[1]];
    if (month !== undefined) {
      return new Date(parseInt(monthDayYear[3], 10), month, parseInt(monthDayYear[2], 10));
    }
  }
  
  // Try DD/MM format (for compiled attendance, assuming current year)
  const ddmm = dateStr.match(/(\d{1,2})\/(\d{1,2})$/);
  if (ddmm) {
    const day = parseInt(ddmm[1], 10);
    const month = parseInt(ddmm[2], 10) - 1;
    return new Date(2025, month, day);
  }
  
  return null;
}

/**
 * Format date as DD/MM for display
 */
export function formatDateShort(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  return `${day}/${month}`;
}

/**
 * Format date as DD/MM/YYYY
 */
export function formatDateFull(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Extract time from clock-in/out string that may contain additional info
 * e.g., "08:12:00 (+07:00) (Mobile phone)" -> "08:12"
 */
export function extractTime(timeStr: string | null | undefined): string | null {
  if (!timeStr || timeStr === '' || timeStr === '__' || timeStr.includes('__')) return null;
  
  // Look for time pattern HH:MM or HH:MM:SS
  const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (timeMatch) {
    return `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}`;
  }
  
  return null;
}

/**
 * Compare two times and return the earlier one
 */
export function getEarlierTime(time1: string | null, time2: string | null): string | null {
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);
  
  if (min1 === null && min2 === null) return null;
  if (min1 === null) return time2;
  if (min2 === null) return time1;
  
  return min1 <= min2 ? time1 : time2;
}

/**
 * Compare two times and return the later one
 */
export function getLaterTime(time1: string | null, time2: string | null): string | null {
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);
  
  if (min1 === null && min2 === null) return null;
  if (min1 === null) return time2;
  if (min2 === null) return time1;
  
  return min1 >= min2 ? time1 : time2;
}
