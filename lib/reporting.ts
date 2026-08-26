import type { ReportingConfig, WorkRecord } from "./models";

function parseActivityDate(activityDate: string) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(activityDate);
  if (!match) throw new Error("Activity date must use YYYY-MM-DD.");
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) throw new Error("Activity date is not a real calendar date.");
  return { year, month };
}

export function deriveSchoolYear(activityDate: string, config: ReportingConfig) {
  const { year, month } = parseActivityDate(activityDate);
  const startYear = month >= config.schoolYearStartMonth ? year : year - 1;
  return `${startYear}-${startYear + 1}`;
}

export function deriveReportingQuarter(activityDate: string, config: ReportingConfig) {
  const { month } = parseActivityDate(activityDate);
  const quarter = config.quarters.find(({ startMonth, endMonth }) => startMonth <= endMonth ? month >= startMonth && month <= endMonth : month >= startMonth || month <= endMonth);
  if (!quarter) throw new Error("No reporting quarter is configured for this date.");
  return quarter.code;
}

export function deriveReportingDays(minutes: number, config: ReportingConfig) {
  if (!Number.isFinite(minutes) || minutes < 0 || config.minutesPerReportingDay <= 0) throw new Error("Reporting minutes and minutes per day must be valid non-negative values.");
  return minutes / config.minutesPerReportingDay;
}

export const deriveStemPocMinutes = (record: WorkRecord) => record.orbit.reportable ? record.orbit.stemPocMinutes : 0;
export const deriveTacMinutes = (record: WorkRecord) => record.orbit.reportable ? record.orbit.tacMinutes : 0;
