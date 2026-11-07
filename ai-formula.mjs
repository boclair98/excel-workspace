const SAFE_FUNCTIONS = new Set(
  "ABS ACOS ACOSH AND ASIN ASINH ATAN ATAN2 ATANH AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS CEILING CEILING.MATH CHAR CHOOSE CLEAN CODE COLUMN COLUMNS CONCAT CONCATENATE COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DATE DATEDIF DATEVALUE DAY DAYS EDATE EOMONTH EXACT EXP FILTER FIND FLOOR FLOOR.MATH HLOOKUP HOUR IF IFERROR IFNA IFS INDEX INT ISBLANK ISERROR ISNA ISNUMBER ISTEXT LARGE LEFT LEN LET LN LOG LOG10 LOOKUP LOWER MATCH MAX MAXIFS MEDIAN MID MIN MINIFS MINUTE MOD MONTH MROUND NOT NOW OR PERCENTILE.INC PERCENTILE.EXC PI POWER PRODUCT PROPER RANK RANK.EQ REPLACE REPT RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SECOND SEQUENCE SIGN SMALL SORT SQRT STDEV STDEV.S STDEV.P SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT TEXT TEXTJOIN TIME TIMEVALUE TODAY TRIM TRUE FALSE UNIQUE UPPER VALUE VLOOKUP WEEKDAY WEEKNUM WORKDAY XLOOKUP XMATCH YEAR YEARFRAC".split(
    " ",
  ),
);
export function safeFormula(value) {
  if (
    typeof value !== "string" ||
    value.length < 2 ||
    value.length > 2001 ||
    !value.startsWith("=") ||
    /[\r\n\u0000-\u001f]/.test(value)
  )
    return null;
  // Quoted strings and sheet names are data. Inspect executable formula tokens separately.
  const tokens = value
    .replace(/"(?:[^"]|"")*"/g, '""')
    .replace(/'(?:[^']|'')*'/g, "''");
  if (/[\[\]\\|{}]/.test(tokens)) return null;
  const names = [...tokens.matchAll(/([A-Z_][A-Z0-9_.]*)\s*\(/gi)].map(
    (match) => match[1].toUpperCase(),
  );
  if (names.some((name) => !SAFE_FUNCTIONS.has(name))) return null;
  // External workbook paths may be enclosed in quoted sheet names.
  if (/\[[^\]]*\]|(?:https?|file|ftp):|\\\\/i.test(value)) return null;
  return value;
}
