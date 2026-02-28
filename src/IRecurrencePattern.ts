import type { EndType, MonthNthSpecific, PatternType, RecurFrequency, WeekSpecific } from "./RecurrencePattern.class.js";

export interface IRecurrencePattern {
  recurFrequency: RecurFrequency;
  patternType: PatternType;
  firstDateTime: Date;
  period: number;
  patternTypeSpecific: null | number | WeekSpecific | MonthNthSpecific;
  endType: EndType;
  occurrenceCount: number;
  firstDOW: number;
  startDate: Date;
  endDate: Date;

  toJSON(): any;
}