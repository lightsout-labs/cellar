import { afterEach, describe, expect, it, vi } from 'vitest';
import { calculate, column, error, expectExcel } from '../../../tests/formula';
import type { Value } from '../value';
import { functions, parse, print } from './index';

type Example = [formula: string, expected: Value];

/** Checks each formula of a table of Excel's examples over the cells. */
function examples(
  title: string,
  cells: Record<string, Value>,
  table: Example[],
  rounded = false,
): void {
  it.each(table)(`${title}: %s gives %j`, (formula, expected) => {
    expectExcel(calculate(formula, cells), expected, rounded);
  });
}

// As in functions.test.ts, each table holds the examples of Excel's
// documentation of the functions, with their data. Dates in cells are the
// serial numbers Excel stores for them.

describe('text functions', () => {
  examples(
    'parts of text',
    {
      ...column('A2', ['Sale Price', 'Sweden', 'Stock Number']),
      B2: 'Fluid Flow',
      ...column('C2', ['Phoenix, AZ', '', '     One   ']),
    },
    [
      ['LEFT(A2,4)', 'Sale'],
      ['LEFT(A3)', 'S'],
      ['RIGHT(A2,5)', 'Price'],
      ['RIGHT(A4)', 'r'],
      ['MID(B2,1,5)', 'Fluid'],
      ['MID(B2,7,20)', 'Flow'],
      ['MID(B2,20,5)', ''],
      ['LEN(C2)', 11],
      ['LEN(C3)', 0],
      ['LEN(C4)', 11],
      ['LEN(1/3)', 17],
    ],
  );

  examples(
    'FIND and SEARCH',
    {
      ...column('A2', ['Miriam McGovern', 'Statements', 'Profit Margin']),
      A5: 'margin',
      A6: 'The "boss" is here.',
      ...column('B2', [
        'Ceramic Insulators #124-TD45-87',
        'Copper Coils #12-671-6772',
        'Variable Resistors #116010',
      ]),
    },
    [
      ['FIND("M",A2)', 1],
      ['FIND("m",A2)', 6],
      ['FIND("M",A2,3)', 8],
      ['MID(B2,1,FIND(" #",B2,1)-1)', 'Ceramic Insulators'],
      ['MID(B3,1,FIND(" #",B3,1)-1)', 'Copper Coils'],
      ['MID(B4,1,FIND(" #",B4,1)-1)', 'Variable Resistors'],
      ['SEARCH("e",A3,6)', 7],
      ['SEARCH(A5,A4)', 8],
      ['REPLACE(A4,SEARCH(A5,A4),6,"Amount")', 'Profit Amount'],
      ['MID(A4,SEARCH(" ",A4)+1,4)', 'Marg'],
      ['SEARCH("""",A6)', 5],
      [
        'MID(A6,SEARCH("""",A6)+1,SEARCH("""",A6,SEARCH("""",A6)+1)-SEARCH("""",A6)-1)',
        'boss',
      ],
      ['SEARCH("m?r*n",A2)', 1],
      ['SEARCH("~?","a?b")', 2],
    ],
  );

  examples(
    'changing text',
    {
      ...column('A2', ['abcdefghijk', '2009', '123456']),
      ...column('B2', ['Sales Data', 'Quarter 1, 2008', 'Quarter 1, 2011']),
      ...column('C2', ['total', 'Yield', 'E. E. Cummings', 'Apt. 2B']),
      ...column('D2', ['this is a TITLE', '2-way street', '76BudGet']),
    },
    [
      ['REPLACE(A2,6,5,"*")', 'abcde*k'],
      ['REPLACE(A3,3,2,"10")', '2010'],
      ['REPLACE(A4,1,3,"@")', '@456'],
      ['SUBSTITUTE(B2,"Sales","Cost")', 'Cost Data'],
      ['SUBSTITUTE(B3,"1","2",1)', 'Quarter 2, 2008'],
      ['SUBSTITUTE(B4,"1","2",3)', 'Quarter 1, 2012'],
      ['REPT("*-",3)', '*-*-*-'],
      ['REPT("-",10)', '----------'],
      ['TRIM(" First   Quarter Earnings ")', 'First Quarter Earnings'],
      ['CLEAN(CHAR(9)&"Monthly report"&CHAR(10))', 'Monthly report'],
      ['UPPER(C2)', 'TOTAL'],
      ['UPPER(C3)', 'YIELD'],
      ['LOWER(C4)', 'e. e. cummings'],
      ['LOWER(C5)', 'apt. 2b'],
      ['PROPER(D2)', 'This Is A Title'],
      ['PROPER(D3)', '2-Way Street'],
      ['PROPER(D4)', '76Budget'],
    ],
  );

  examples(
    'characters and comparison',
    { ...column('A2', ['word', 'Word', 'w ord']), ...column('B2', ['word']) },
    [
      ['CHAR(65)', 'A'],
      ['CHAR(33)', '!'],
      ['CODE("A")', 65],
      ['CODE("!")', 33],
      ['CHAR(128)', '€'],
      ['CODE("€")', 128],
      ['CODE("é")', 233],
      ['CODE("中")', 63],
      ['UNICHAR(66)', 'B'],
      ['UNICHAR(32)', ' '],
      ['UNICODE(" ")', 32],
      ['UNICODE("B")', 66],
      ['UNICODE(UNICHAR(128512))', 128512],
      ['EXACT(A2,B2)', true],
      ['EXACT(A3,B2)', false],
      ['EXACT(A4,B2)', false],
    ],
  );

  examples(
    'joining text',
    {
      ...column('A2', ['brook trout', 'species', 32]),
      ...column('B2', ['a1', 'a2', 'a4', 'a5', 'a6', 'a7']),
      ...column('C2', ['b1', 'b2', 'b4', 'b5', 'b6', 'b7']),
      ...column('D2', [
        'US Dollar',
        'Australian Dollar',
        null,
        'Chinese Yuan',
        'Hong Kong Dollar',
      ]),
    },
    [
      [
        'CONCATENATE("Stream population for ",A2," ",A3," is ",A4,"/mile.")',
        'Stream population for brook trout species is 32/mile.',
      ],
      [
        'CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")',
        'The sun will come up tomorrow.',
      ],
      ['CONCAT(B2:C7)', 'a1b1a2b2a4b4a5b5a6b6a7b7'],
      [
        'TEXTJOIN(", ",TRUE,D2:D6)',
        'US Dollar, Australian Dollar, Chinese Yuan, Hong Kong Dollar',
      ],
      ['TEXTJOIN(",",FALSE,D2:D4)', 'US Dollar,Australian Dollar,'],
      ['TEXTJOIN("-",TRUE,"a","",1,TRUE)', 'a-1-TRUE'],
    ],
  );

  examples(
    'numbers and text',
    {
      ...column('A2', ['Rainfall', 19, true]),
      ...column('B2', [1234.567, -1234.567, 44.332, -0.123, 99.888]),
    },
    [
      ['T(A2)', 'Rainfall'],
      ['T(A3)', ''],
      ['T(A4)', ''],
      ['VALUE("$1,000")', 1000],
      ['VALUE("16:48:00")-VALUE("12:00:00")', 0.2],
      ['NUMBERVALUE("2.500,27",",",".")', 2500.27],
      ['NUMBERVALUE("3.5%")', 0.035],
      ['NUMBERVALUE(" 1 000 ")', 1000],
      ['TEXT(1234.567,"$#,##0.00")', '$1,234.57'],
      ['TEXT(0.285,"0.0%")', '28.5%'],
      ['TEXT(4.34,"# ?/?")', '4 1/3'],
      ['TEXT(12200000,"0.00E+00")', '1.22E+07'],
      ['TEXT(1234,"0000000")', '0001234'],
      ['TEXT(DATE(2008,1,1),"dddd, mmm d, yyyy")', 'Tuesday, Jan 1, 2008'],
      ['TEXT("abc","0.00")', 'abc'],
      ['FIXED(B2,1)', '1,234.6'],
      ['FIXED(B2,-1)', '1,230'],
      ['FIXED(B3,-1,TRUE)', '-1230'],
      ['FIXED(B4)', '44.33'],
      ['DOLLAR(B2,2)', '$1,234.57'],
      ['DOLLAR(B2,-2)', '$1,200'],
      ['DOLLAR(B3,-2)', '($1,200)'],
      ['DOLLAR(B5,4)', '($0.1230)'],
      ['DOLLAR(B6)', '$99.89'],
    ],
  );
});

describe('date and time functions', () => {
  examples(
    'dates from parts and parts of dates',
    {
      ...column('A2', [2011, 1, 1]),
      ...column('B2', [39634, 40364, 39492]),
    },
    [
      ['DATE(A2,A3,A4)', 40544],
      ['DATE(2008,14,2)', 39846],
      ['DATE(2008,-3,2)', 39327],
      ['DATE(2008,1,35)', 39482],
      ['DATE(2008,1,-15)', 39432],
      ['DATE(108,1,2)', 39449],
      // Excel takes 1900 as a leap year, as Lotus 1-2-3 did.
      ['DATE(1900,2,29)', 60],
      ['DATE(1900,3,1)', 61],
      ['DAY(60)', 29],
      ['MONTH(60)', 2],
      ['DAY(0)', 0],
      ['YEAR(0)', 1900],
      ['YEAR(B2)', 2008],
      ['YEAR(B3)', 2010],
      ['MONTH("15-Apr-11")', 4],
      ['DAY("15-Apr-11")', 15],
      ['YEAR(2958465)', 9999],
      ['WEEKDAY(B4)', 5],
      ['WEEKDAY(B4,2)', 4],
      ['WEEKDAY(B4,3)', 3],
      ['WEEKDAY(1)', 1],
      ['WEEKDAY(B4,16)', 6],
    ],
  );

  examples(
    'times',
    { ...column('A2', [0.75, '7/18/2011 7:45', '4/21/2012']) },
    [
      ['HOUR(A2)', 18],
      ['HOUR(A3)', 7],
      ['HOUR(A4)', 0],
      ['MINUTE("12:45:00 PM")', 45],
      ['SECOND("4:48:18 PM")', 18],
      ['SECOND("4:48 PM")', 0],
      ['TIME(12,0,0)', 0.5],
      ['TIME(16,48,10)', 0.7001157],
      ['TIME(27,0,0)', 0.125],
      ['TIMEVALUE("2:24 AM")', 0.1],
      ['TIMEVALUE("22-Aug-2011 6:35 AM")', 0.274305556],
      ['DATEVALUE("8/22/2011")', 40777],
      ['DATEVALUE("22-MAY-2011")', 40685],
      ['DATEVALUE("2011/02/23")', 40597],
      ['DATEVALUE("22-Aug-2011 6:35 AM")', 40777],
    ],
    true,
  );

  examples(
    'spans of time',
    {
      ...column('A2', [44561, 44197]),
      ...column('B2', [40544, 40573, 40575, 40908]),
      ...column('C2', [40558, 40544]),
      ...column('D2', [40909, 41120]),
      ...column('E2', [36892, 37622, 37043, 37483]),
    },
    [
      ['DAYS("15-MAR-2021","1-FEB-2021")', 42],
      ['DAYS(A2,A3)', 364],
      ['DAYS360(B2,B3)', 29],
      ['DAYS360(B2,B4)', 30],
      ['DAYS360(B2,B5)', 360],
      ['DAYS360(DATE(2011,1,31),DATE(2011,3,31),TRUE)', 60],
      ['EDATE(C2,1)', 40589],
      ['EDATE(C2,-1)', 40527],
      ['EDATE(C2,2)', 40617],
      ['EDATE(DATE(2011,1,31),1)', 40602],
      ['EOMONTH(C3,1)', 40602],
      ['EOMONTH(C3,-3)', 40482],
      ['YEARFRAC(D2,D3)', 0.58055556],
      ['YEARFRAC(D2,D3,1)', 0.57650273],
      ['YEARFRAC(D2,D3,3)', 0.57808219],
      ['YEARFRAC(D2,D3,2)', 0.58611111],
      // Over a 29 February, and over several years by their average length.
      ['YEARFRAC(DATE(2011,6,1),DATE(2012,5,1),1)', 0.9153005],
      ['YEARFRAC(DATE(2012,7,30),DATE(2010,1,1),1)', 2.5757299],
      ['YEARFRAC(D2,D3,4)', 0.58055556],
      ['DATEDIF(E2,E3,"Y")', 2],
      ['DATEDIF(E4,E5,"D")', 440],
      ['DATEDIF(E4,E5,"YD")', 75],
      ['DATEDIF(E4,E5,"MD")', 14],
      ['DATEDIF(E4,E5,"M")', 14],
      ['DATEDIF(E4,E5,"YM")', 2],
    ],
    true,
  );

  examples(
    'weeks and working days',
    {
      ...column('A2', [41183, 41334, 41235, 41247, 41295]),
      ...column('B2', [39722, 151, 39778, 39786, 39834]),
      C2: 40977,
    },
    [
      ['NETWORKDAYS(A2,A3)', 110],
      ['NETWORKDAYS(A2,A3,A4)', 109],
      ['NETWORKDAYS(A2,A3,A4:A6)', 107],
      ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))', 22],
      ['NETWORKDAYS(DATE(2006,1,6),DATE(2006,1,2))', -5],
      ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7)', 24],
      ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001")', 22],
      ['WORKDAY(B2,B3)', 39933],
      ['WORKDAY(B2,B3,B4:B6)', 39938],
      ['WORKDAY(B2,-5)', 39715],
      ['WORKDAY.INTL(DATE(2012,1,1),90,11)', 41013],
      ['WORKDAY.INTL(B2,B3,,B4:B6)', 39938],
      ['TEXT(WORKDAY.INTL(DATE(2012,1,1),30,17),"m/dd/yyyy")', '2/05/2012'],
      ['WEEKNUM(C2)', 10],
      ['WEEKNUM(C2,2)', 11],
      ['WEEKNUM(DATE(2012,1,1),21)', 52],
      ['ISOWEEKNUM(C2)', 10],
      ['ISOWEEKNUM(DATE(2021,1,3))', 53],
    ],
  );

  describe('the clock', () => {
    afterEach(() => vi.useRealTimers());

    it('gives TODAY and NOW in the local time', () => {
      vi.useFakeTimers({ now: new Date(2011, 11, 30, 18, 0) });
      expect(calculate('TODAY()')).toBe(40907);
      expect(calculate('NOW()')).toBeCloseTo(40907.75, 9);
    });
  });
});

describe('lookup and reference functions', () => {
  const densities = {
    ...column('A2', [0.457, 0.525, 0.606, 0.675, 0.746, 0.835, 0.946, 1.09]),
    A10: 1.29,
    ...column('B2', [3.55, 3.25, 2.93, 2.75, 2.57, 2.38, 2.17, 1.95, 1.71]),
    ...column('C2', [500, 400, 300, 250, 200, 150, 100, 50, 0]),
  };

  examples('VLOOKUP', densities, [
    ['VLOOKUP(1,A2:C10,2)', 2.17],
    ['VLOOKUP(1,A2:C10,3,TRUE)', 100],
    ['VLOOKUP(0.7,A2:C10,3,FALSE)', error('#N/A')],
    ['VLOOKUP(0.1,A2:C10,2,TRUE)', error('#N/A')],
    ['VLOOKUP(2,A2:C10,2,TRUE)', 1.71],
    ['VLOOKUP(0.746,A2:C10,3,FALSE)', 200],
    ['VLOOKUP(1,A:C,3)', 100],
  ]);

  examples(
    'HLOOKUP',
    {
      ...column('A1', ['Axles', 4, 5, 6]),
      ...column('B1', ['Bearings', 4, 7, 8]),
      ...column('C1', ['Bolts', 9, 10, 11]),
    },
    [
      ['HLOOKUP("Axles",A1:C4,2,TRUE)', 4],
      ['HLOOKUP("Bearings",A1:C4,3,FALSE)', 7],
      ['HLOOKUP("B",A1:C4,3,TRUE)', 5],
      ['HLOOKUP("Bolts",A1:C4,4)', 11],
      ['HLOOKUP("b*",A1:C4,2,FALSE)', 4],
    ],
  );

  examples(
    'MATCH, XMATCH and XLOOKUP',
    {
      ...column('A2', ['Bananas', 'Oranges', 'Apples', 'Pears']),
      ...column('B2', [25, 38, 40, 41]),
      ...column('C2', [0, 9701, 39476, 84201, 160726, 204101]),
      ...column('D2', [0.1, 0.22, 0.24, 0.32, 0.35, 0.37]),
      ...column('E2', ['China', 'India', 'United States', 'India']),
      ...column('F2', ['+86', '+91', '+1', '+0']),
    },
    [
      ['MATCH(39,B2:B5,1)', 2],
      ['MATCH(41,B2:B5,0)', 4],
      ['MATCH(40,B2:B5,-1)', error('#N/A')],
      ['MATCH("app*",A2:A5,0)', 3],
      ['MATCH(100,B2:B5)', 4],
      ['XMATCH("Pears",A2:A5)', 4],
      ['XMATCH(39,B2:B5,1)', 3],
      ['XMATCH(39,B2:B5,-1)', 2],
      ['XMATCH("?ears",A2:A5,2)', 4],
      ['XMATCH("India",E2:E5,0,-1)', 4],
      ['XLOOKUP("India",E2:E5,F2:F5)', '+91'],
      ['XLOOKUP("India",E2:E5,F2:F5,,0,-1)', '+0'],
      ['XLOOKUP("India",E2:E5,F2:F5,,,)', '+91'],
      ['XMATCH(39,B2:B5,,)', error('#N/A')],
      ['XLOOKUP("Brazil",E2:E5,F2:F5)', error('#N/A')],
      ['XLOOKUP("Brazil",E2:E5,F2:F5,"Not found")', 'Not found'],
      ['XLOOKUP(46523,C2:C7,D2:D7,0,-1)', 0.24],
      ['XLOOKUP(46523,C2:C7,D2:D7,0,1)', 0.32],
      ['XLOOKUP(46523,C2:C7,D2:D7,0,1,2)', 0.32],
      ['XLOOKUP(39476,C2:C7,D2:D7,0,0,2)', 0.24],
    ],
  );

  examples(
    'LOOKUP',
    {
      ...column('A2', [4.14, 4.19, 5.17, 5.77, 6.39]),
      ...column('B2', ['red', 'orange', 'yellow', 'green', 'blue']),
    },
    [
      ['LOOKUP(4.19,A2:A6,B2:B6)', 'orange'],
      ['LOOKUP(5.75,A2:A6,B2:B6)', 'yellow'],
      ['LOOKUP(7.66,A2:A6,B2:B6)', 'blue'],
      ['LOOKUP(0,A2:A6,B2:B6)', error('#N/A')],
      ['LOOKUP(5.2,A2:A6)', 5.17],
      ['LOOKUP(5.2,A2:B6)', 'yellow'],
    ],
  );

  examples(
    'INDEX and CHOOSE',
    {
      ...column('A2', ['Apples', 'Bananas']),
      ...column('B2', ['Lemons', 'Pears']),
      ...column('C2', [1, 2]),
      ...column('D2', ['1st', '2nd', '3rd', 'Finished']),
      ...column('E2', ['Nails', 'Screws', 'Nuts', 'Bolts']),
    },
    [
      ['INDEX(A2:B3,2,2)', 'Pears'],
      ['INDEX(A2:B3,2,1)', 'Bananas'],
      ['INDEX(A2:A3,2)', 'Bananas'],
      ['INDEX(A2:B2,2)', 'Lemons'],
      ['SUM(INDEX(A2:C3,0,3))', 3],
      ['SUM(INDEX(C2:C3,1,1,1))', 1],
      ['INDEX(7,1,1)', 7],
      ['CHOOSE(2,D2,D3,D4,D5)', '2nd'],
      ['CHOOSE(4,E2,E3,E4,E5)', 'Bolts'],
      ['CHOOSE(3,"Wide",115,"world",8)', 'world'],
      ['SUM(CHOOSE(2,A1:A9,C2:C3))', 3],
      ['CHOOSE(1,"a",1/0)', 'a'],
    ],
  );

  examples('positions and addresses', {}, [
    ['ROW()', 100],
    ['ROW(C10)', 10],
    ['ROW(C10:D20)', 10],
    ['COLUMN()', 26],
    ['COLUMN(D10)', 4],
    ['ROWS(C1:E4)', 4],
    ['COLUMNS(C1:E4)', 3],
    ['ROWS(1)', 1],
    ['COLUMNS(A:B)', 2],
    ['ADDRESS(2,3)', '$C$2'],
    ['ADDRESS(2,3,2)', 'C$2'],
    ['ADDRESS(2,3,3)', '$C2'],
    ['ADDRESS(2,3,4)', 'C2'],
    ['ADDRESS(2,3,2,FALSE)', 'R2C[3]'],
    ['ADDRESS(2,3,,FALSE)', 'R2C3'],
    ['ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")', "'[Book1]Sheet1'!R2C3"],
    ['ADDRESS(2,3,1,FALSE,"EXCEL SHEET")', "'EXCEL SHEET'!R2C3"],
    ['ADDRESS(2,28,1,TRUE,"Data")', 'Data!$AB$2'],
  ]);
});

describe('information functions', () => {
  examples(
    'the IS functions',
    {
      ...column('A2', ['Gold', 'Region1', error('#REF!'), 330.92]),
      A6: error('#N/A'),
    },
    [
      ['ISLOGICAL(TRUE)', true],
      ['ISLOGICAL("TRUE")', false],
      ['ISNUMBER(4)', true],
      ['ISREF(G8)', true],
      ['ISREF(XYZ1)', false],
      ['ISREF(1)', false],
      ['ISBLANK(A2)', false],
      ['ISBLANK(A9)', true],
      ['ISERROR(A4)', true],
      ['ISNA(A4)', false],
      ['ISNA(A6)', true],
      ['ISERR(A6)', false],
      ['ISERR(A4)', true],
      ['ISNUMBER(A5)', true],
      ['ISTEXT(A3)', true],
      ['ISNONTEXT(A3)', false],
      ['ISNONTEXT(A9)', true],
      ['ISEVEN(-1)', false],
      ['ISEVEN(2.5)', true],
      ['ISEVEN(5)', false],
      ['ISEVEN(0)', true],
      ['ISODD(-1)', true],
      ['ISODD(2.5)', false],
      ['ISODD(5)', true],
    ],
  );

  examples(
    'values and errors',
    {
      ...column('A2', [7, 'Even', true, 40650]),
      B2: 'Smith',
      C2: error('#NUM!'),
    },
    [
      ['N(A2)', 7],
      ['N(A3)', 0],
      ['N(A4)', 1],
      ['N(A5)', 40650],
      ['N("7")', 0],
      ['NA()', error('#N/A')],
      ['TYPE(B2)', 2],
      ['TYPE("Mr. "&B2)', 2],
      ['TYPE(2+B2)', 16],
      ['TYPE(A2)', 1],
      ['TYPE(A4)', 4],
      ['ERROR.TYPE(#NULL!)', 1],
      ['ERROR.TYPE(1/0)', 2],
      ['ERROR.TYPE(C2)', 6],
      ['ERROR.TYPE(#N/A)', 7],
      ['ERROR.TYPE(1)', error('#N/A')],
      [
        'IF(ERROR.TYPE(C2)<3,CHOOSE(ERROR.TYPE(C2),"Ranges do not intersect","The divisor is zero"))',
        false,
      ],
    ],
  );
});

describe('errors of the text, date, lookup and information functions', () => {
  const cells = {
    ...column('A1', [1, 2, 3]),
    ...column('B1', ['a', 'b', 'c']),
  };

  it.each<[string, string]>([
    ['MID("abc",0,1)', '#VALUE!'],
    ['LEFT("abc",-1)', '#VALUE!'],
    ['FIND("z","abc")', '#VALUE!'],
    ['FIND("a","abc",4)', '#VALUE!'],
    ['SEARCH("a","abc",0)', '#VALUE!'],
    ['CHAR(0)', '#VALUE!'],
    ['CHAR(256)', '#VALUE!'],
    ['CODE("")', '#VALUE!'],
    ['UNICHAR(0)', '#VALUE!'],
    ['REPT("ab",-1)', '#VALUE!'],
    ['SUBSTITUTE("a","a","b",0)', '#VALUE!'],
    ['VALUE("abc")', '#VALUE!'],
    ['UPPER(1/0)', '#DIV/0!'],
    ['CONCAT(A1:A3,#N/A)', '#N/A'],
    ['DATE(-1,1,1)', '#NUM!'],
    ['DATE(10000,1,1)', '#NUM!'],
    ['DAY(-1)', '#NUM!'],
    ['YEAR(2958466)', '#NUM!'],
    ['TIME(-1,0,0)', '#NUM!'],
    ['DATEVALUE("abc")', '#VALUE!'],
    ['DATEVALUE(40000)', '#VALUE!'],
    ['TIMEVALUE("abc")', '#VALUE!'],
    ['WEEKDAY(1,4)', '#NUM!'],
    ['WEEKNUM(1,3)', '#NUM!'],
    ['YEARFRAC(1,2,5)', '#NUM!'],
    ['DATEDIF(2,1,"D")', '#NUM!'],
    ['DATEDIF(1,2,"W")', '#NUM!'],
    ['NETWORKDAYS.INTL(1,9,"1111111")', '#VALUE!'],
    ['NETWORKDAYS.INTL(1,9,8)', '#NUM!'],
    ['VLOOKUP(1,A1:B3,3,FALSE)', '#REF!'],
    ['VLOOKUP(1,A1:B3,0,FALSE)', '#VALUE!'],
    ['HLOOKUP(1,A1:B3,4,FALSE)', '#REF!'],
    ['MATCH(9,A1:A3,0)', '#N/A'],
    ['MATCH(1,A1:B3,0)', '#N/A'],
    ['INDEX(A1:B3,4,1)', '#REF!'],
    ['INDEX(A1:B3,1,1,2)', '#REF!'],
    ['INDEX(A1:B3,-1,1)', '#VALUE!'],
    ['CHOOSE(3,1,2)', '#VALUE!'],
    ['CHOOSE(0,1,2)', '#VALUE!'],
    ['XLOOKUP(1,A1:A3,B1:B2)', '#VALUE!'],
    ['XMATCH(1,A1:A3,3)', '#VALUE!'],
    ['ADDRESS(0,1)', '#VALUE!'],
    ['ISEVEN("a")', '#VALUE!'],
    ['N(1/0)', '#DIV/0!'],
  ])('%s gives %s', (formula, code) => {
    expect(calculate(formula, cells)).toEqual(error(code));
  });

  it.each([
    'LEFT()',
    'MID("a",1)',
    'DATE(1,2)',
    'TODAY(1)',
    'VLOOKUP(1,A1:B2)',
    'NA(1)',
    'CHOOSE(1)',
    'ROW(1)',
  ])('refuses %s, with the wrong number of arguments', (formula) => {
    expect(parse(formula)).toBeNull();
  });
});

describe('the library', () => {
  it('has at least 150 functions', () => {
    expect(functions.length).toBeGreaterThanOrEqual(150);
  });

  it('writes the prefix for files for the newer functions of these families', () => {
    const ast = parse(
      'CONCAT(TEXTJOIN(",",TRUE,A1),XLOOKUP(1,A1:A2,B1:B2),DAYS(2,1),VLOOKUP(1,A1:B2,2))',
    )!;
    expect(print(ast, { file: true })).toBe(
      '_xlfn.CONCAT(_xlfn.TEXTJOIN(",",TRUE,A1),_xlfn.XLOOKUP(1,A1:A2,B1:B2),_xlfn.DAYS(2,1),VLOOKUP(1,A1:B2,2))',
    );
  });
});
