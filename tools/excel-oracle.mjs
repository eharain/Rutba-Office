// Asks Excel itself what a few hundred formulas and number formats come to,
// and writes the answers to tests/fixtures/excel-oracle.json. The test
// (tests/excel-oracle.test.js) holds the engine to them, on any machine.
//
//   node tools/excel-oracle.mjs       needs Excel; nothing is saved but the answers
//
// A case's answer is what Excel showed, never what this code thinks it should be.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

/** The cells every formula can see. A leading ' is text; = a formula; null blank. */
export const SHEETS = {
  Sheet1: {
    A1: 1, A2: 2, A3: 3, A4: 'text', A5: 5, A6: true, A8: 0, A9: -4, A10: 2.5,
    B1: 10, B2: 'a', B3: 'b', B4: 40, B5: '=""', B6: "'12", B7: 7.5, B9: '=""', B10: 100,
    C1: 'apple', C2: 'Apricot', C3: 'banana', C4: 'cherry', C6: 'a*c', C7: 'x?y',
    E1: 1, E2: 2, E3: 3, E4: 4, E5: 5, E6: 6,
    F1: 10, F2: 20, F3: '=NA()', F4: 40, F5: '=1/0', F6: 60,
    H1: 46000, H2: 46031, H3: 46059,
    M1: true, M2: "'TRUE", M3: 46037, M4: "'5", M5: 5,
  },
  'Дані': { A1: 7, A2: 8 },
  'Q1 Sales': { A1: 11, A2: 22 },
};

export const FORMULAS = [
  '=DATE(99,1,1)', '=DATE(0,1,1)', '=DATE(-1,1,1)', '=DATE(10000,1,1)', '=DATE(9999,12,31)', '=DATE(1899,12,31)', '=DATE(1900,1,1)', '=DATE(1900,1,0)', '=DATE(1900,2,29)', '=DATE(1900,3,0)', '=DATE(2024.9,1.9,1.9)', '=DATE(2024,14,1)', '=DATE(2024,1,-1)', '=DATE(1900,1,-1)',
  // What a blank is, and what empty text is.
  '=A7', '=IF(TRUE,A7)', '=A7&""', '=INDEX(A1:A10,7)', '=IF(A7,1,2)', '=B5&"x"', '=LEN(A7)', '=SUM(B5)', '=AVERAGE(B5,1)', '=N(B5)', '=ABS(B5)',
  '=MAX(A6,1)', '=MIN(A6)', '=COUNT(A6)', '=COUNT(A4)', '=SUM(A4)', '=PRODUCT(A4)', '=AVERAGE(A4)', '=AND(A6)', '=AND(B5)', '=NOT(B5)', '=OR(A4)',
  '=TRUE+1', '=A6+1', '=A7=""', '=A7=0', '=A7=FALSE', '=""=0', '=B5=""', '=B5=A7', '=MATCH("",A7:A8,0)', '=VLOOKUP(0,A7:B8,2,FALSE)',
  // Text that is a number, a date or a time when arithmetic asks.
  '="2026-01-15"+1', '="15 Jan 2026"+0', '="12:30"+0', '="50%"+0', '="1,234"+1', '=VALUE("15 Jan 2026")',
  // What criteria match.
  '=COUNTIF(M1:M5,"TRUE")', '=COUNTIF(M1:M5,"<>5")', '=COUNTIF(M1:M5,"2026-01-15")', '=COUNTIF(M1:M5,">=2026-01-15")', '=COUNTIF(M1:M5,5)', '=COUNTIF(M1:M5,"=5")', '=COUNTIF(M1:M5,M5)', '=COUNTIF(M1:M5,A7)',
  '=COUNTIF(C1:C7,"<banana")', '=COUNTIF(C1:C7,">=banana")', '=COUNTIF(C1:C7,"<>a*")', '=COUNTIF(C1:C7,"a?ple")', '=SUMIFS(E1:E6,E1:E6,">2",E1:E6,"<5")', '=COUNTIFS(E1:E6,">2",F1:F6,">30")', '=AVERAGEIF(E1:E6,">4")',
  // INDEX with a row or column of 0, and on a single row.
  '=SUM(INDEX(E1:F6,0,2))', '=SUM(INDEX(E1:F2,1,0))', '=INDEX({1,2,3},2)', '=INDEX(E1:E6,2)', '=HLOOKUP(20,E1:F2,1,FALSE)',
  // Operators, in Excel's own order.
  '=-2^2', '=2^3^2', '=-2^-2', '=2^-1', '=--2', '=1+-2*3', '=2*-3^2', '=-A2^2', '=(-2)^2', '=10%^2', '=2^3%', '=-2%', '=50%%',
  '=1-2-3', '=2/3/4', '="a"&1+2', '=1&2+3', '=1=1=TRUE', '=1<2=TRUE', '=+3', '=+"x"', '=-"3"', '=2^0.5', '=0^0', '=(-8)^(1/3)',
  '=1/0', '=0/0', '="a"+1', '="3"+1', '="3"*"4"', '=TRUE+1', '=TRUE&""', '=1=1', '="a"="A"', '="a"<"b"', '=1<"a"', '="a"<TRUE', '=TRUE>"z"',
  '=A8=FALSE', '=A1=TRUE', '=SUM({-1,2})', '=INDEX({-1,2;3,-4},2,2)', '=SUM({1,2;3,4})', '=ROWS({1,2;3,4})', '=COLUMNS({1,2;3,4})', '={1,2}+1',
  // Sheet names.
  '=Дані!A1+1', '=SUM(Дані!A1:A2)', "='Q1 Sales'!A2", "=SUM('Q1 Sales'!A1:A2)", '=sheet1!A1', '=SHEET1!E2', "='sheet1'!E3",
  // What an error does, and what it does not.
  '=ISNUMBER(MATCH(99,A1:A3,0))', '=ISNA(MATCH(99,A1:A3,0))', '=IFERROR(1/0,"x")', '=IFNA(NA(),"x")', '=IFNA(1/0,"x")', '=ISERROR(NA())', '=ISERR(NA())', '=ISERR(1/0)',
  '=COUNT(F1:F6)', '=COUNTA(F1:F6)', '=VLOOKUP(2,E1:F6,2,FALSE)', '=VLOOKUP(3,E1:F6,2,FALSE)', '=VLOOKUP(6,E1:F6,2,FALSE)', '=SUM(F1:F6)', '=SUM(F1:F2)', '=AVERAGE(F1:F6)', '=MAX(F1:F6)',
  '=COUNTIF(F1:F6,">15")', '=COUNTIF(F1:F6,"#N/A")', '=SUMIF(F1:F6,">15")', '=SUMIF(E1:E6,">3",F1:F6)', '=SUMIF(E1:E2,">0",F1:F2)', '=IF(TRUE,1,NA())', '=IF(FALSE,1/0,2)', '=IF(NA(),1,2)',
  '=AND(TRUE,NA())', '=OR(TRUE,NA())', '=CHOOSE(1,2,NA())', '=CHOOSE(2,2,NA())', '=INDEX(F1:F6,1)', '=INDEX(F1:F6,3)', '=MATCH(40,F1:F6,0)', '=MATCH(60,F1:F6,0)', '=ISTEXT(F3)', '=ISNUMBER(F3)', '=ISBLANK(F3)', '=ISLOGICAL(F3)',
  '=TYPE(NA())', '=TYPE(1)', '=TYPE("a")', '=TYPE(TRUE)', '=TYPE({1,2})', '=ERROR.TYPE(NA())', '=ERROR.TYPE(1/0)', '=N(NA())', '=N("a")', '=N(TRUE)', '=N(5)', '=T(5)', '=T("a")', '=F3+1', '=F3&"x"', '=NOT(NA())', '=ROWS(F1:F6)', '=COLUMNS(E1:F6)', '=SUMPRODUCT(E1:E2,F1:F2)', '=SUMPRODUCT(E1:E6,F1:F6)',
  '=LOOKUP(2,E1:E6,F1:F6)', '=LOOKUP(3.5,E1:E6)', '=LOOKUP(0,E1:E6,F1:F6)', '=HLOOKUP(1,E1:F2,2,FALSE)', '=AGGREGATE(9,6,F1:F6)', '=AGGREGATE(4,6,F1:F6)', '=AGGREGATE(1,6,F1:F2)', '=AGGREGATE(14,6,F1:F6,1)',
  // Blank and empty text are not the same.
  '=A7=""', '=A7=0', '=ISBLANK(A7)', '=ISBLANK(B5)', '=B5=""', '=B5=0', '=A7=B5', '=A7&"x"', '=A7+1', '=B5+1', '=B5&"x"', '=LEN(A7)', '=LEN(B5)', '=COUNTBLANK(A1:B10)', '=COUNTA(A1:B10)', '=COUNT(A1:B10)',
  '=COUNTIF(A1:B10,"")', '=COUNTIF(A1:B10,"<5")', '=COUNTIF(A1:B10,">5")', '=COUNTIF(A1:B10,"<>")', '=COUNTIF(A1:B10,"=")', '=COUNTIF(A1:B10,"<>1")', '=COUNTIF(A1:B10,5)', '=COUNTIF(A1:B10,"5")', '=COUNTIF(A1:B10,"12")', '=COUNTIF(A1:B10,TRUE)', '=COUNTIF(A1:B10,"text")', '=COUNTIF(A1:B10,"TEXT")',
  '=COUNTIF(C1:C7,"a*")', '=COUNTIF(C1:C7,"?pple")', '=COUNTIF(C1:C7,"a~*c")', '=COUNTIF(C1:C7,"x~?y")', '=COUNTIF(C1:C7,"<>a*")', '=COUNTIF(C1:C7,">b")', '=COUNTIF(C1:C7,"*")', '=SUMIF(A1:A10,">2")', '=SUMIF(A1:A10,"<2")', '=AVERAGEIF(A1:A10,">0")', '=COUNTIFS(A1:A3,">1",E1:E3,"<3")', '=SUMIFS(F1:F2,E1:E2,">1")', '=MAX(A7,1)', '=MIN(A7:A8)', '=SUM(A1:A10)', '=AVERAGE(A1:A10)', '=MAX(A1:A10)', '=MIN(A1:A10)', '=SUM(A6)', '=SUM(TRUE,1)', '=SUM("3",1)', '=AVERAGE(A7:A7)', '=PRODUCT(A1:A3)', '=COUNT(TRUE,"1","a")', '=COUNTA(A7,B5)',
  // Functions.
  '=SIN(1)', '=COS(1)', '=TAN(1)', '=ASIN(0.5)', '=ACOS(0.5)', '=ATAN(1)', '=ATAN2(1,2)', '=PI()', '=DEGREES(PI())', '=RADIANS(180)', '=SINH(1)', '=COSH(1)', '=TANH(1)', '=EXP(1)', '=LN(10)', '=LOG(100)', '=LOG(8,2)', '=LOG10(1000)', '=SQRT(16)', '=SQRT(-1)', '=ABS(-3)', '=SIGN(-3)', '=FACT(5)', '=GCD(12,18)', '=LCM(4,6)', '=COMBIN(5,2)', '=PERMUT(5,2)',
  '=ROUND(2.675,2)', '=ROUND(1.005,2)', '=ROUND(-2.5,0)', '=ROUND(1234.5,-2)', '=ROUNDUP(1.01,1)', '=ROUNDDOWN(-1.99,1)', '=TRUNC(4.35,2)', '=TRUNC(-4.9)', '=TRUNC(1234.56,-2)', '=INT(-1.5)', '=INT(1.5)', '=MOD(-3,2)', '=MOD(3,-2)', '=MOD(5,0)', '=MOD(5.5,2)', '=QUOTIENT(7,2)', '=FLOOR(0.3,0.1)', '=FLOOR(7,2)', '=FLOOR(-7,-2)', '=FLOOR(-7,2)', '=FLOOR(7,-2)', '=CEILING(0.3,0.1)', '=CEILING(7,2)', '=CEILING(-7,-2)', '=CEILING(-7,2)', '=MROUND(7,2)', '=ODD(2)', '=EVEN(3)', '=POWER(2,10)', '=FLOOR.MATH(-7,2)', '=CEILING.MATH(-7,2)',
  '=PERCENTILE(A1:A3,0.5)', '=PERCENTILE(E1:E6,0.25)', '=PERCENTILE.INC(E1:E6,0.9)', '=PERCENTILE.EXC(E1:E6,0.5)', '=QUARTILE(E1:E6,1)', '=QUARTILE.INC(E1:E6,3)', '=MEDIAN(E1:E6)', '=MODE(1,2,2,3)', '=STDEV(E1:E6)', '=STDEV.S(E1:E6)', '=STDEV.P(E1:E6)', '=VAR(E1:E6)', '=VAR.P(E1:E6)', '=LARGE(E1:E6,2)', '=SMALL(E1:E6,2)', '=RANK(3,E1:E6)', '=RANK(3,E1:E6,1)', '=CORREL(E1:E3,H1:H3)', '=SLOPE(H1:H3,E1:E3)', '=INTERCEPT(H1:H3,E1:E3)',
  '=SUBSTITUTE("aXbXc","X","-")', '=SUBSTITUTE("aXbXc","X","-",2)', '=SUBSTITUTE("aXbXc","X","-",9)', '=REPLACE("abcdef",2,3,"Z")', '=LEFT("abc")', '=LEFT("abc",2)', '=RIGHT("abc",2)', '=MID("abcdef",2,3)', '=LEN("abc")', '=UPPER("aB")', '=LOWER("aB")', '=PROPER("hello wORLD")', '=TRIM("  a   b  ")', '=CLEAN("a"&CHAR(7)&"b")', '=FIND("b","abcb")', '=FIND("b","abcb",3)', '=FIND("z","abc")', '=SEARCH("B","abc")', '=SEARCH("a?c","xabc")', '=EXACT("a","A")', '=REPT("ab",2)', '=CONCATENATE("a",1,TRUE)', '=CONCAT("a",1)', '=TEXTJOIN(",",TRUE,"a","","b")', '=TEXTJOIN(",",FALSE,"a","","b")', '=VALUE("12")', '=VALUE("1,234.5")', '=VALUE("abc")', '=VALUE("12%")', '=CHAR(65)', '=CODE("A")', '=UNICHAR(1046)', '=UNICODE("Ж")', '=FIXED(1234.567,1)', '=FIXED(1234.567,1,TRUE)',
  '=TEXT(1234.5,"#,##0.00")', '=TEXT(0.5,"0%")', '=TEXT(46000,"yyyy-mm-dd")', '=TEXT(46000,"dddd")', '=TEXT(46000,"mmm d")', '=TEXT(1.5,"0.0")', '=TEXT(5,"000")', '=TEXT(-5,"0;(0)")', '=TEXT("a","@")', '=TEXT(0.75,"h:mm")', '=TEXT(1234,"0.0E+00")',
  '=DATE(2026,1,31)', '=DATE(2026,13,1)', '=DATE(2026,1,0)', '=YEAR(46000)', '=MONTH(46000)', '=DAY(46000)', '=WEEKDAY(46000)', '=WEEKDAY(46000,2)', '=WEEKNUM(46000)', '=ISOWEEKNUM(46000)', '=HOUR(0.75)', '=MINUTE(0.76)', '=SECOND(0.123)', '=TIME(12,30,0)', '=EDATE(46000,1)', '=EDATE(46000,-13)', '=EOMONTH(46000,0)', '=EOMONTH(46000,1)', '=DATEVALUE("2026-01-15")', '=DATEVALUE("15 Jan 2026")', '=TIMEVALUE("12:30")', '=DAYS(46031,46000)', '=DAYS360(46000,46031)', '=YEARFRAC(46000,46031)', '=DATEDIF(46000,46400,"m")', '=DATEDIF(46000,46400,"d")', '=NETWORKDAYS(46000,46031)', '=NETWORKDAYS(46000,46031,H2:H2)', '=WORKDAY(46000,5)', '=WORKDAY(46000,-5)', '=NETWORKDAYS.INTL(46000,46031,1)', '=WORKDAY.INTL(46000,5,1)',
  '=IF(1,2,3)', '=IF(0,2,3)', '=IF("a",2,3)', '=IF(A7,2,3)', '=IF(1,2)', '=IF(0,2)', '=IFS(FALSE,1,TRUE,2)', '=SWITCH(2,1,"a",2,"b")', '=AND(1,0)', '=OR(0,0)', '=NOT(0)', '=XOR(TRUE,FALSE)', '=TRUE()', '=FALSE()', '=AND(A1:A3)', '=AND("a")', '=AND("TRUE")',
  '=ROW(A5)', '=COLUMN(C1)', '=ROWS(A1:A3)', '=ROW()', '=COLUMN()', '=ADDRESS(1,2)', '=ADDRESS(1,2,4)', '=INDEX(A1:B2,2,2)', '=MATCH("b",B1:B4,0)', '=MATCH(2,E1:E6)', '=MATCH(2,E1:E6,1)', '=MATCH(9,E1:E6,1)', '=MATCH("A*",C1:C4,0)', '=VLOOKUP("banana",C1:C4,1,FALSE)', '=VLOOKUP("APPLE",C1:C4,1,FALSE)', '=VLOOKUP(2.5,E1:F6,2)', '=VLOOKUP(2,E1:F6,3,FALSE)', '=XLOOKUP(2,E1:E6,F1:F6)', '=XLOOKUP(9,E1:E6,F1:F6,"none")', '=XMATCH(2,E1:E6)',
  '=SUM(INDIRECT("A1:A3"))', '=INDIRECT("A2")', '=INDIRECT("Дані!A1")', "=INDIRECT(\"'Q1 Sales'!A1\")", '=SUM(OFFSET(A1,1,0,2,1))', '=OFFSET(A1,1,0)', '=INDIRECT("R1C1",FALSE)', '=INDIRECT("nothing")',
  '=PMT(0.05/12,60,10000)', '=FV(0.05,10,-100)', '=PV(0.05,10,-100)', '=NPV(0.1,100,200)', '=IRR({-100,60,60})', '=RATE(10,-100,800)', '=NPER(0.05,-100,800)', '=ISEVEN(2)', '=ISODD(2)', '=DEC2HEX(255)', '=HEX2DEC("FF")', '=BIN2DEC("101")', '=DEC2BIN(5)', '=ROMAN(14)', '=ARABIC("XIV")', '=BASE(255,16)', '=DECIMAL("FF",16)', '=SUMSQ(1,2)', '=SUMX2MY2({1,2},{1,1})', '=MAXA(A1:A6)', '=MINA(A1:A6)', '=AVERAGEA(A1:A6)', '=COUNTIF(A1:A3,A2)', '=NA()', '=SUM()+0',
];

/** A value and the format it wears, and what Excel prints. */
export const FORMATS = [
  ['abc', '"x"@'], ['abc', '0;"n";"z";@" y"'], ['abc', '0'], [123456789012345678, '0'], [0.5, '.00'], [1234.5, '# ##0'], [1234.5, '0.0"x"0'], [0.999999, 'h:mm:ss'], [1.5, '[h]:mm'], [1.5, '[ss]'], [-5, '[<0]0;0'], [-5, '0;0'], [5, '[Red]0'], [5, '"[x]"0'], [-0.001, '0.00'], [1234, '#,'], [1234567, '0,'], [46000, 'y'],
  [1234.5, '[$£-809]#,##0.00'], [1234.5, '[$€-407] #,##0.00'], [1234.5, '[$$-409]#,##0'], [1500, '[>=1000]#,##0;0'], [15, '[>=1000]#,##0;0'], [15, '[<10]"small";[>=100]"big";"mid"'], [5, '[<10]"small";[>=100]"big";"mid"'], [500, '[<10]"small";[>=100]"big";"mid"'],
  [123456789, '000-00-0000'], [5551234, '###-####'], [2.675, '0.00'], [1.005, '0.00'], [0.125, '0.00'], [2.5, '0'], [3.5, '0'], [-2.5, '0'], [1234.5678, 'General'], [5, 'General" units"'], [1234.5678, '0.00" kg"'], [1234.5, '#,##0'], [1234567, '#,##0,'], [1234567, '#,##0.0,,"M"'], [0.1234, '0%'], [0.1234, '0.00%'], [1234, '0.0E+00'], [0.00012, '0.00E+00'], [0, '0.00;-0.00;"zero"'], [-3, '0;(0)'], [-3, '0;[Red]0'], [3, '[Blue]0'], [0.5, '# ?/?'], [1.25, '# ?/?'], [1.333, '# ??/??'], [0.5, '?/8'], [12, '@'], [12, '"x"0"y"'], [12, '\\a0'], [1.5, '0.?'], [1.5, '0.0#'], [1, '0.##'], [0, '#'], [0, '#.##'], [12.3, '000.000'], [1234.5, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'], [-1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'], [0, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  [46000.75, 'yyyy-mm-dd'], [46000.75, 'dd/mm/yy'], [46000.75, 'm/d/yyyy'], [46000.75, 'mmm'], [46000.75, 'mmmm'], [46000.75, 'mmmmm'], [46000.75, 'ddd'], [46000.75, 'dddd'], [46000.75, 'd-mmm-yy'], [46000.75, 'h:mm'], [46000.75, 'h:mm AM/PM'], [46000.75, 'hh:mm:ss'], [46000.75, 'yyyy-mm-dd hh:mm'], [0.5, '[h]:mm'], [1.5, '[h]:mm'], [1.5, '[mm]:ss'], [0.51, 'mm:ss'], [0.123456, 'h:mm:ss.00'], [46000, 'yy'], [46000, 'yyy'], [60, 'yyyy-mm-dd'], [61, 'yyyy-mm-dd'], [1, 'yyyy-mm-dd'],
];

/** Asked again of the same book on the 1904 date system. */
export const ON_1904 = {
  formulas: ['=DATE(2024,1,15)', '=YEAR(0)', '=MONTH(0)', '=DAY(0)', '=YEAR(1000)', '=WEEKDAY(0)', '=WEEKDAY(1000)', '=DATEVALUE("2024-03-01")', '=EDATE(DATE(2024,1,31),1)', '=EOMONTH(DATE(2024,2,10),0)', '=DAYS(DATE(2024,3,1),DATE(2024,1,1))', '=TEXT(1000,"yyyy-mm-dd")', '=TEXT(0,"yyyy-mm-dd dddd")', '=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', '=WORKDAY(DATE(2024,1,1),10)', '=DATE(1904,1,1)', '=DATE(1904,1,2)', '=DATE(1903,1,1)', '=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))', '=WEEKNUM(DATE(2024,3,1))', '=DATEDIF(DATE(2020,1,1),DATE(2024,3,1),"m")'],
  formats: [[0, 'yyyy-mm-dd'], [1000, 'yyyy-mm-dd'], [46000.75, 'dd/mm/yy hh:mm'], [1000, 'dddd mmmm d']],
};

function main() {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-excel-oracle-'));
  const input = path.join(dir, 'in.json');
  const output = path.join(dir, 'out.json');
  fs.writeFileSync(input, JSON.stringify({ sheets: SHEETS, formulas: FORMULAS, formats: FORMATS, on1904: ON_1904 }));
  const script = path.join(dir, 'ask.ps1');
  fs.writeFileSync(script, `${String.fromCharCode(0xfeff)}$ErrorActionPreference = 'Stop'
$in = Get-Content -Raw -Encoding UTF8 '${input}' | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
function Answer($c) {
  $v = $c.Value2
  if ($v -is [int] -and $c.Text -like '#*') { return @{ error = [string]$c.Text } }
  if ($v -is [bool]) { return @{ value = $v } }
  if ($v -is [double]) { return @{ value = $v } }
  if ($null -eq $v) { return @{ value = $null } }
  return @{ value = [string]$v }
}
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
  $first = $true
  foreach ($p in $in.sheets.PSObject.Properties) {
    if ($first) { $ws = $wb.Worksheets.Item(1); $first = $false } else { $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws.Name = $p.Name
    foreach ($cell in $p.Value.PSObject.Properties) {
      $c = $ws.Range($cell.Name)
      if ($cell.Value -is [string]) { $c.Formula = $cell.Value } elseif ($cell.Value -is [bool]) { $c.Value2 = [bool]$cell.Value } else { $c.Value2 = [double]$cell.Value }
    }
  }
  # On the sheet with the data, clear of it: a formula's A1 means that sheet's A1.
  $ws = $wb.Worksheets.Item('Sheet1')
  $out = @{ formulas = @(); formats = @() }
  foreach ($f in $in.formulas) {
    $c = $ws.Range('J20'); $c.Clear() | Out-Null
    try {
      $c.Formula2 = $f
      $xl.Calculate()
      $out.formulas += ,(@{ formula = $f; answer = (Answer $c) })
    } catch { $out.formulas += ,(@{ formula = $f; refused = $true }) }
  }
  foreach ($pair in $in.formats) {
    $c = $ws.Range('L20'); $c.Clear() | Out-Null
    $c.ColumnWidth = 80
    $c.NumberFormat = $pair[1]; if ($pair[0] -is [string]) { $c.Value2 = $pair[0] } else { $c.Value2 = [double]$pair[0] }
    $out.formats += ,(@{ value = $pair[0]; format = $pair[1]; text = [string]$c.Text })
  }
  $wb.Date1904 = $true
  $out.on1904 = @{ formulas = @(); formats = @() }
  foreach ($f in $in.on1904.formulas) {
    $c = $ws.Range('J20'); $c.Clear() | Out-Null
    $c.Formula2 = $f
    $xl.Calculate()
    $out.on1904.formulas += ,(@{ formula = $f; answer = (Answer $c) })
  }
  foreach ($pair in $in.on1904.formats) {
    $c = $ws.Range('L20'); $c.Clear() | Out-Null
    $c.ColumnWidth = 80
    $c.NumberFormat = $pair[1]; $c.Value2 = [double]$pair[0]
    $out.on1904.formats += ,(@{ value = $pair[0]; format = $pair[1]; text = [string]$c.Text })
  }
  $out.excel = [string]$xl.Version + '.' + [string]$xl.Build
  $wb.Close($false)
  [IO.File]::WriteAllText('${output}', ($out | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
} finally { $xl.Quit() }
`.replace(/\n/g, '\r\n'));
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], { encoding: 'utf8', timeout: 400000 });
  if (r.status !== 0) throw new Error(`Excel did not answer: ${r.stderr || r.stdout || r.error}`);
  const answers = JSON.parse(fs.readFileSync(output, 'utf8'));
  const fixture = path.join(root, 'tests', 'fixtures', 'excel-oracle.json');
  fs.writeFileSync(fixture, `${JSON.stringify({ sheets: SHEETS, ...answers }, null, 1)}\n`);
  fs.rmSync(dir, { recursive: true, force: true });
  console.log(`excel-oracle: ${answers.formulas.length} formulas and ${answers.formats.length} formats answered by Excel ${answers.excel}`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
