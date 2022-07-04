import * as coda from "@codahq/packs-sdk";

import { Options, RRule, rrulestr } from 'rrule'
import { getEnumValues } from "./utils/enum.helpers";
import { convertFrequencyStringToRRuleConstant, convertUserWeekdayStringToRRuleConstant, convertUserMonthStringToRRuleConstant, convertUserByWeekdayToRRuleConstant, CouldNotConvertError } from "./rrule_converter";
import { MONTHS, WEEKDAYS } from "./utils/date.constants";
import { deleteUndefinedProps } from "./utils/object.helpers";
import { capitalize, escapeControlCodes, unescapeControlCodes } from "./utils/string.helpers";

const MAX_CONSUMED_RECURRENCES = 1000;
const CACHE_TTL_SECS = 0;
const BLANK = "";

export const pack = coda.newPack();

// note that the abbreviations "MO, TU, WE" etc. or the day-of-week as a number in the ISO week numbering system, where Monday is 1 can be used anywhere in lieu of their complete counterparts Monday, Tuesday, Wednesday. These arguments are also not case-sensitive, where possible.
// note that the abbreviations "JAN, FEB, MAR" etc. or the month numbers in the ISO system, where January is 1, can be ''

pack.addFormula({
  name: "CreateRRule",
  description: "Returns a Recurrence Rule String",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    coda.makeParameter({
      name: "frequency",
      description: `One of the following: ${RRule.FREQUENCIES.map(capitalize).join(", ")}.`,
      type: coda.ParameterType.String,
      autocomplete: RRule.FREQUENCIES.map(frequencyString => ({
        display: capitalize(frequencyString),
        value: frequencyString
      }))
    }),
    coda.makeParameter({
      name: "dtstart",
      description: "The recurrence start. If not given, the current date and time will be used instead.",
      type: coda.ParameterType.Date,
      optional: true,
    }),
    coda.makeParameter({
      name: "interval",
      description: 'The interval between each frequency iteration. For example, when using "Yearly" frequency, an interval of 2 means once every two years, but with "Hourly" frequency, it means once every two hours. The default interval is 1.',
      type: coda.ParameterType.Number,
      optional: true
    }),
    coda.makeParameter({
      name: "wkst",
      description: `The week start day specifying the first day of the week. Must be one of ${getEnumValues(WEEKDAYS).map(capitalize).join(", ")}. This will affect recurrences based on weekly periods. The default week start is Monday.`,
      type: coda.ParameterType.String,
      autocomplete: getEnumValues(WEEKDAYS).map(weekdayString => ({
        display: capitalize(weekdayString),
        value: weekdayString
      })),
      optional: true,
    }),
    coda.makeParameter({
      name: "count",
      description: "How many occurrences will be generated. The default is to recur infinitely.",
      type: coda.ParameterType.Number,
      optional: true,
      suggestedValue: 13, // based on Google Calendar, arbitrarily
    }),
    coda.makeParameter({
      name: "until",
      description: "Specifies the limit of the recurrence. If a recurrence instance happens to be the same as the date given in the `until` argument, this will be the last occurrence.",
      type: coda.ParameterType.Date,
      optional: true,
    }),
    coda.makeParameter({
      name: "timezone",
      description: "Specifies the TZID parameter in the [RFC](https://tools.ietf.org/html/rfc5545#section-3.2.19) with an IANA string recognized by the [Intl API](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl).",
      type: coda.ParameterType.String,
      // autocomplete: Intl.supportedValuesOf('timeZone'),
      optional: true,
    }),
    coda.makeParameter({
      name: "bysetpos",
      description: 'A list of positive or negative numbers, each of which will specify an occurrence number, corresponding to the nth occurrence of the rule inside the frequency period. For example, a `bysetpos` of -1 if combined with a "Monthly" frequency, and a `byweekday` of "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", will result in the last work day of every month.',
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "bymonth",
      description: `A list of months, meaning the months to apply the recurrence to. Must be some of ${getEnumValues(MONTHS).map(capitalize).join(", ")}.`,
      type: coda.ParameterType.SparseStringArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "bymonthday",
      description: "A list of numbers, meaning the month days to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byyearday",
      description: "A list of numbers, meaning the year days to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byweekno",
      description: "A list of numbers, meaning the week numbers (in the ISO week numbering system, where week 1 contains the first Thursday of the year and weeks start on Monday) to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byweekday",
      description: `A list of weekdays specifying the weekdays where the recurrence will be applied. Must be some of ${getEnumValues(WEEKDAYS).map(capitalize).join(", ")}. It's also possible to prefix the weekdays with a number n for the weekday instances, which will mean the nth occurrence of this weekday in the period. For example, with "Monthly" frequency, using "1Friday" or "-1Friday" in \`byweekday\` will specify the first or last Friday of the month where the recurrence happens. Notice that the RFC documentation, this is specified as BYDAY, but was renamed to avoid the ambiguity of that argument.`,
      type: coda.ParameterType.SparseStringArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byhour",
      description: "A list of numbers, meaning the hours to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byminute",
      description: "A list of numbers, meaning the minutes to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "bysecond",
      description: "A list of numbers, meaning the seconds to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
  ],

  resultType: coda.ValueType.String,
  execute: async function ([frequency, dtstart, interval, wkst, count, until, timezone, bysetpos, bymonth, bymonthday, byyearday, byweekno, byweekday, byhour, byminute, bysecond], context) {
    let options: Partial<Options>;

    try {
      options = deleteUndefinedProps({
        freq: convertFrequencyStringToRRuleConstant(frequency),
        dtstart,
        interval,
        wkst: wkst !== undefined ? convertUserWeekdayStringToRRuleConstant(wkst) : undefined,
        count,
        until,
        tzid: timezone,
        bysetpos,
        bymonth: bymonth !== undefined ? bymonth.map(convertUserMonthStringToRRuleConstant) : undefined,
        bymonthday,
        byyearday,
        byweekno,
        byweekday: byweekday !== undefined ? byweekday.map(convertUserByWeekdayToRRuleConstant) : undefined,
        byhour,
        byminute,
        bysecond,
      });
    } catch (e) {
      if (e instanceof CouldNotConvertError) {
        throw new coda.UserVisibleError(e.message);
      } else {
        throw e;
      }
    }

    const rule = new RRule(options);

    return escapeControlCodes(rule.toString());
  },
});

const rruleParameter = coda.makeParameter({
  name: "rrule",
  description: "The RRule string to operate on.",
  type: coda.ParameterType.String
});

pack.addFormula({
  name: "ModifyRRule",
  description: "Returns the Recurrence Rule String constructed from overwriting the given rrule options with the new parameters",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter,
    coda.makeParameter({
      name: "frequency",
      description: `One of the following: ${RRule.FREQUENCIES.map(capitalize).join(", ")}.`,
      type: coda.ParameterType.String,
      autocomplete: RRule.FREQUENCIES.map(frequencyString => ({
        display: capitalize(frequencyString),
        value: frequencyString
      })),
      optional: true
    }),
    coda.makeParameter({
      name: "dtstart",
      description: "The recurrence start. If not given, the current date and time will be used instead.",
      type: coda.ParameterType.Date,
      optional: true,
    }),
    coda.makeParameter({
      name: "interval",
      description: 'The interval between each frequency iteration. For example, when using "Yearly" frequency, an interval of 2 means once every two years, but with "Hourly" frequency, it means once every two hours. The default interval is 1.',
      type: coda.ParameterType.Number,
      optional: true
    }),
    coda.makeParameter({
      name: "wkst",
      description: `The week start day specifying the first day of the week. Must be one of ${getEnumValues(WEEKDAYS).map(capitalize).join(", ")}. This will affect recurrences based on weekly periods. The default week start is Monday.`,
      type: coda.ParameterType.String,
      autocomplete: getEnumValues(WEEKDAYS).map(weekdayString => ({
        display: capitalize(weekdayString),
        value: weekdayString
      })),
      optional: true,
    }),
    coda.makeParameter({
      name: "count",
      description: "How many occurrences will be generated. The default is to recur infinitely.",
      type: coda.ParameterType.Number,
      optional: true,
      suggestedValue: 13, // based on Google Calendar, arbitrarily
    }),
    coda.makeParameter({
      name: "until",
      description: "Specifies the limit of the recurrence. If a recurrence instance happens to be the same as the date given in the `until` argument, this will be the last occurrence.",
      type: coda.ParameterType.Date,
      optional: true,
    }),
    coda.makeParameter({
      name: "timezone",
      description: "Specifies the TZID parameter in the [RFC](https://tools.ietf.org/html/rfc5545#section-3.2.19) with an IANA string recognized by the [Intl API](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl).",
      type: coda.ParameterType.String,
      // autocomplete: Intl.supportedValuesOf('timeZone'),
      optional: true,
    }),
    coda.makeParameter({
      name: "bysetpos",
      description: 'A list of positive or negative numbers, each of which will specify an occurrence number, corresponding to the nth occurrence of the rule inside the frequency period. For example, a `bysetpos` of -1 if combined with a "Monthly" frequency, and a `byweekday` of "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", will result in the last work day of every month.',
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "bymonth",
      description: `A list of months, meaning the months to apply the recurrence to. Must be some of ${getEnumValues(MONTHS).map(capitalize).join(", ")}.`,
      type: coda.ParameterType.SparseStringArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "bymonthday",
      description: "A list of numbers, meaning the month days to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byyearday",
      description: "A list of numbers, meaning the year days to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byweekno",
      description: "A list of numbers, meaning the week numbers (in the ISO week numbering system, where week 1 contains the first Thursday of the year and weeks start on Monday) to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byweekday",
      description: `A list of weekdays specifying the weekdays where the recurrence will be applied. Must be some of ${getEnumValues(WEEKDAYS).map(capitalize).join(", ")}. It's also possible to prefix the weekdays with a number n for the weekday instances, which will mean the nth occurrence of this weekday in the period. For example, with "Monthly" frequency, using "1Friday" or "-1Friday" in \`byweekday\` will specify the first or last Friday of the month where the recurrence happens. Notice that the RFC documentation, this is specified as BYDAY, but was renamed to avoid the ambiguity of that argument.`,
      type: coda.ParameterType.SparseStringArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byhour",
      description: "A list of numbers, meaning the hours to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "byminute",
      description: "A list of numbers, meaning the minutes to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
    coda.makeParameter({
      name: "bysecond",
      description: "A list of numbers, meaning the seconds to apply the recurrence to.",
      type: coda.ParameterType.SparseNumberArray,
      optional: true,
    }),
  ],

  resultType: coda.ValueType.String,
  execute: async function ([rrule, frequency, dtstart, interval, wkst, count, until, timezone, bysetpos, bymonth, bymonthday, byyearday, byweekno, byweekday, byhour, byminute, bysecond], context) {
    let options: Partial<Options>;
    const unescapedRRule = unescapeControlCodes(rrule);
    try {
      options = rrulestr(unescapedRRule).origOptions;
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    let newOptions: Partial<Options>;

    try {
      newOptions = deleteUndefinedProps({
        freq: frequency !== undefined ? convertFrequencyStringToRRuleConstant(frequency) : undefined,
        dtstart,
        interval,
        wkst: wkst !== undefined ? convertUserWeekdayStringToRRuleConstant(wkst) : undefined,
        count,
        until,
        tzid: timezone,
        bysetpos,
        bymonth: bymonth !== undefined ? bymonth.map(convertUserMonthStringToRRuleConstant) : undefined,
        bymonthday,
        byyearday,
        byweekno,
        byweekday: byweekday !== undefined ? byweekday.map(convertUserByWeekdayToRRuleConstant) : undefined,
        byhour,
        byminute,
        bysecond,
      });
    } catch (e) {
      if (e instanceof CouldNotConvertError) {
        throw new coda.UserVisibleError(e.message);
      } else {
        throw e;
      }
    }

    options = {
      ...options,
      ...newOptions,
    }

    const rule = new RRule(options);

    return escapeControlCodes(rule.toString());
  },
});

const limitingIterator = (limit: number) => (date: Date, i: number) => i < limit;

const limitParameter = coda.makeParameter({
  name: "limit",
  description: `The limit on the the number of occurrences to return. The default is 100. The maximum is ${MAX_CONSUMED_RECURRENCES}.`,
  type: coda.ParameterType.Number,
  suggestedValue: 100,
  optional: true,
});

pack.addFormula({
  name: "All",
  description: "Returns the first `limit` occurrences of the RRule.",

  parameters: [
    rruleParameter,
    limitParameter
  ],

  resultType: coda.ValueType.Array,
  items: { type: coda.ValueType.String, codaType: coda.ValueHintType.Date },

  execute: async function([rrule, limit = 100], context) {
    limit = Math.min(limit, MAX_CONSUMED_RECURRENCES); // silently cap
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try{
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    return rule.all(limitingIterator(limit)).map(date => date.toISOString());
  }
});

pack.addFormula({
  name: "Between",
  description: "Returns the first `limit` occurrences of the RRule between `after` and `before`.",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter,
    coda.makeParameter({
      name: "after",
      description: "The lower bound on the occurrences.",
      type: coda.ParameterType.Date
    }),
    coda.makeParameter({
      name: "before",
      description: "The upper bound on the occurrences.",
      type: coda.ParameterType.Date,
    }),
    limitParameter,
    coda.makeParameter({
      name: "include",
      description: "Defines what happens if `after` and/or `before` are themselves occurrences. If `true`, they will be included in the list, if they are found in the recurrence set. The default is `false`.",
      type: coda.ParameterType.Boolean,
      suggestedValue: false,
      optional: true
    }),
  ],

  resultType: coda.ValueType.Array,
  items: { type: coda.ValueType.String, codaType: coda.ValueHintType.DateTime },

  execute: async function ([rrule, after, before, limit = 100, include = false], context) {
    limit = Math.min(limit, MAX_CONSUMED_RECURRENCES); // silently cap
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try {
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    return rule.between(after, before, include, limitingIterator(limit)).map(date => date.toISOString());

  }
});

pack.addFormula({
  name: "First",
  description: "Returns the first recurrence of the RRule",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter
  ],

  resultType: coda.ValueType.String,
  codaType: coda.ValueHintType.DateTime,

  execute: async function([rrule], context) {
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try {
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    return rule.all().length > 0 ? rule.all()[0].toISOString() : BLANK;
  }
});

const dateParameter = coda.makeParameter({
  name: "date",
  description: "The date threshold.",
  type: coda.ParameterType.Date
});

const includeParameter = coda.makeParameter({
  name: "include",
  description: "Defines what happens if `date` is an occurrence. If `true`, if `date` itself is an occurrence, it will be returned. The default is `false`.",
  type: coda.ParameterType.Boolean,
  suggestedValue: false,
  optional: true,
});

pack.addFormula({
  name: "After",
  description: "Returns the first recurrence after the given date",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter,
    dateParameter,
    includeParameter
  ],

  resultType: coda.ValueType.String,
  codaType: coda.ValueHintType.DateTime,

  execute: async function([rrule, date, include = false], context) {
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try {
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    let result: Date;
    return (result = rule.after(date, include)) ? result.toISOString() : BLANK;
  }
});

pack.addFormula({
  name: "Before",
  description: "Returns the last recurrence before the given date",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter,
    dateParameter,
    includeParameter
  ],

  resultType: coda.ValueType.String,
  codaType: coda.ValueHintType.DateTime,

  execute: async function([rrule, date, include = false], context) {
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try {
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    let result: Date;
    return (result = rule.before(date, include)) ? result.toISOString() : BLANK;
  }
});

pack.addFormula({
  name: "ToText",
  description: "Returns a textual representation of the given RRule",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter
  ],

  resultType: coda.ValueType.String,

  execute: async function([rrule], context) {
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try {
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    return rule.toText();
  }
});

pack.addFormula({
  name: "IsFullyConvertibleToText",
  description: "Provides a hint on whether all the options the RRule has are convertible to text",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    rruleParameter
  ],

  resultType: coda.ValueType.Boolean,

  execute: async function([rrule], context) {
    const unescapedRRule = unescapeControlCodes(rrule);
    let rule: RRule;
    try {
      rule = rrulestr(unescapedRRule);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    return rule.isFullyConvertibleToText();
  }
});

const naturalLanguageTextExamples = [
  "Every day",
  "Every week",
  "Every month",
  "Every weekday",
  "Every 2 weeks on Tuesday",
  "Every week on Monday, Wednesday",
  "Every month on the 2nd last Friday for 7 times",
  "Every 6 months"
]

pack.addFormula({
  name: "FromText",
  description: "Returns a RRule constructed from the parsed text",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [
    coda.makeParameter({
      name: "text",
      description: `The natural language text to parse into a RRule. Examples: ${naturalLanguageTextExamples.join(" | ")}`,
      type: coda.ParameterType.String,
      autocomplete: naturalLanguageTextExamples
    })
  ],

  resultType: coda.ValueType.String,

  execute: async function([text], context) {
    let rule: RRule;
    try {
      rule = RRule.fromText(text);
    } catch (e) {
      throw new coda.UserVisibleError(e.message);
    }
    return escapeControlCodes(rule.toString());
  }
})

pack.addFormula({
  name: "Frequencies",
  description: "Returns the list of supported frequencies",
  cacheTtlSecs: CACHE_TTL_SECS,

  parameters: [],

  resultType: coda.ValueType.Array,
  items: { type: coda.ValueType.String },

  execute: async function([], context) {
    return RRule.FREQUENCIES.map(capitalize);
  }
});
