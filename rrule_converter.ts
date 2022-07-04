import RRule, { ByWeekday, Frequency, Weekday } from "rrule";

class CouldNotConvertError extends Error {
  constructor(...params) {
    // Pass remaining arguments (including vendor specific ones) to parent constructor
    super(...params);

    // Maintains proper stack trace for where our error was thrown (only available on V8)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, CouldNotConvertError);
    }

    this.name = 'CouldNotConvertError';
  }
}

function convertUserWeekdayStringToRRuleConstant(weekday: string): Weekday {
  switch (weekday.toLowerCase().trim()) {
    case "1":
    case "monday":
    case "mo":
      return RRule.MO;
    case "2":
    case "tuesday":
    case "tu":
      return RRule.TU;
    case "3":
    case "wednesday":
    case "we":
      return RRule.WE;
    case "4":
    case "thursday":
    case "th":
      return RRule.TH;
    case "5":
    case "friday":
    case "fr":
      return RRule.FR;
    case "6":
    case "saturday":
    case "sa":
      return RRule.SA;
    case "7":
    case "sunday":
    case "su":
      return RRule.SU;
    default:
      throw new CouldNotConvertError;
  }
}

function convertUserMonthStringToRRuleConstant(month: string): number {
  switch (month.toLowerCase().trim()) {
    case "1":
    case "january":
    case "jan":
      return 1;
    case "2":
    case "february":
    case "feb":
      return 2;
    case "3":
    case "march":
    case "ma":
      return 3;
    case "4":
    case "april":
    case "apr":
      return 4;
    case "5":
    case "may":
      return 5;
    case "6":
    case "june":
    case "jun":
      return 6;
    case "7":
    case "july":
    case "jul":
      return 7;
    case "8":
    case "august":
    case "aug":
      return 8;
    case "9":
    case "september":
    case "sep":
      return 9;
    case "10":
    case "october":
    case "oct":
      return 10;
    case "11":
    case "november":
    case "nov":
      return 11;
    case "12":
    case "december":
    case "dec":
      return 12
    default:
      throw new CouldNotConvertError;
  }
  
}

function convertFrequencyStringToRRuleConstant(frequency: string): Frequency {
  frequency = frequency.trim().toUpperCase();

  const frequencyStringToRRuleConstants = {
    YEARLY: RRule.YEARLY,
    MONTHLY: RRule.MONTHLY,
    WEEKLY: RRule.WEEKLY,
    DAILY: RRule.DAILY,
    HOURLY: RRule.HOURLY,
    MINUTELY: RRule.MINUTELY,
    SECONDLY: RRule.SECONDLY,
  }
  if (frequencyStringToRRuleConstants.hasOwnProperty(frequency)) {
    return frequencyStringToRRuleConstants[frequency]
  } else {
    throw new CouldNotConvertError;
  }
}

function convertUserByWeekdayToRRuleConstant(byWeekday: string): ByWeekday {
  byWeekday = byWeekday.toLowerCase().trim();

  const byWeekdayWordMatch = byWeekday.match(/\w+/);

  if (byWeekdayWordMatch.length != 1) {
    throw new CouldNotConvertError;
  }

  const weekday = byWeekdayWordMatch[0];

  let n: number;
  // handle prefix case
  if (/^[-+]?[0-9]+/.test(byWeekday)) {
    // get prefix
    n = parseInt(byWeekday.match(/^[-+]?[0-9]+/)[0]);
    return convertUserWeekdayStringToRRuleConstant(weekday).nth(n);
  } else {
    return convertUserWeekdayStringToRRuleConstant(weekday);
  }
}

export {
  convertFrequencyStringToRRuleConstant,
  convertUserByWeekdayToRRuleConstant,
  convertUserMonthStringToRRuleConstant,
  convertUserWeekdayStringToRRuleConstant,
  CouldNotConvertError
}