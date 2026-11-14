import { describe, expect, it } from 'vitest';
import { calculate, column, error, expectExcel } from '../../../tests/formula';
import type { Value } from '../value';
import { parse, print } from './index';

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

// Each table holds the examples of Excel's documentation of the
// functions, with the data of those examples where they need cells. A
// table marked `true` has results that the documentation shows rounded.

describe('math and trigonometry functions', () => {
  examples(
    'signs, roots and powers',
    { A2: -4, A3: 19.5 },
    [
      ['ABS(2)', 2],
      ['ABS(-2)', 2],
      ['ABS(A2)', 4],
      ['SIGN(10)', 1],
      ['SIGN(4-4)', 0],
      ['SIGN(-0.00001)', -1],
      ['SQRT(16)', 4],
      ['SQRT(ABS(-16))', 4],
      ['SQRTPI(1)', 1.772454],
      ['SQRTPI(2)', 2.506628],
      ['POWER(5,2)', 25],
      ['POWER(98.6,3.2)', 2401077.222],
      ['POWER(4,5/4)', 5.656854249],
      ['EXP(1)', 2.71828183],
      ['EXP(2)', 7.3890561],
      ['LN(86)', 4.4543473],
      ['LN(2.7182818)', 1],
      ['LN(EXP(3))', 3],
      ['LOG(10)', 1],
      ['LOG(8,2)', 3],
      ['LOG(86,2.7182818)', 4.4543473],
      ['LOG10(86)', 1.9345],
      ['LOG10(10)', 1],
      ['LOG10(1E5)', 5],
      ['LOG10(10^5)', 5],
      ['SUMSQ(3,4)', 25],
      ['PI()', 3.14159265358979],
      ['PI()/2', 1.570796327],
      ['PI()*(3^2)', 28.27433388],
      ['A3-INT(A3)', 0.5],
    ],
    true,
  );

  examples(
    'trigonometry',
    {},
    [
      ['ACOS(-0.5)', 2.094395102],
      ['ACOS(-0.5)*180/PI()', 120],
      ['DEGREES(ACOS(-0.5))', 120],
      ['ACOSH(1)', 0],
      ['ACOSH(10)', 2.9932228],
      ['ASIN(-0.5)', -0.523598776],
      ['DEGREES(ASIN(-0.5))', -30],
      ['ASINH(-2.5)', -1.647231146],
      ['ASINH(10)', 2.99822295],
      ['ATAN(1)', 0.785398163],
      ['ATAN2(1,1)', 0.785398163],
      ['ATAN2(-1,-1)', -2.35619449],
      ['DEGREES(ATAN2(-1,-1))', -135],
      ['ATANH(0.76159416)', 1.00000001],
      ['ATANH(-0.1)', -0.100335348],
      ['COS(1.047)', 0.5001711],
      ['COS(60*PI()/180)', 0.5],
      ['COS(RADIANS(60))', 0.5],
      ['COSH(4)', 27.308233],
      ['COSH(EXP(1))', 7.6101251],
      ['DEGREES(PI())', 180],
      ['RADIANS(270)', 4.712389],
      ['SIN(PI())', 0],
      ['SIN(PI()/2)', 1],
      ['SIN(30*PI()/180)', 0.5],
      ['SIN(RADIANS(30))', 0.5],
      ['SINH(1)', 1.175201194],
      ['SINH(-1)', -1.175201194],
      ['TAN(0.785)', 0.99920399],
      ['TAN(45*PI()/180)', 1],
      ['TAN(RADIANS(45))', 1],
      ['TANH(-2)', -0.964028],
      ['TANH(0)', 0],
      ['TANH(0.5)', 0.462117],
    ],
    true,
  );

  examples('rounding', {}, [
    ['ROUND(2.15,1)', 2.2],
    ['ROUND(2.149,1)', 2.1],
    ['ROUND(-1.475,2)', -1.48],
    ['ROUND(21.5,-1)', 20],
    ['ROUND(626.3,-3)', 1000],
    ['ROUND(1.98,-1)', 0],
    ['ROUND(-50.55,-2)', -100],
    ['ROUND(1.005,2)', 1.01],
    ['ROUNDDOWN(3.2,0)', 3],
    ['ROUNDDOWN(76.9,0)', 76],
    ['ROUNDDOWN(3.14159,3)', 3.141],
    ['ROUNDDOWN(-3.14159,1)', -3.1],
    ['ROUNDDOWN(31415.92654,-2)', 31400],
    ['ROUNDUP(3.2,0)', 4],
    ['ROUNDUP(76.9,0)', 77],
    ['ROUNDUP(3.14159,3)', 3.142],
    ['ROUNDUP(-3.14159,1)', -3.2],
    ['ROUNDUP(31415.92654,-2)', 31500],
    ['ROUNDUP(0.1+0.2,1)', 0.3],
    ['MROUND(10,3)', 9],
    ['MROUND(-10,-3)', -9],
    ['MROUND(1.3,0.2)', 1.4],
    ['INT(8.9)', 8],
    ['INT(-8.9)', -9],
    ['TRUNC(8.9)', 8],
    ['TRUNC(-8.9)', -8],
    ['TRUNC(0.45)', 0],
    ['TRUNC(3.14159,2)', 3.14],
    ['EVEN(1.5)', 2],
    ['EVEN(3)', 4],
    ['EVEN(2)', 2],
    ['EVEN(-1)', -2],
    ['ODD(1.5)', 3],
    ['ODD(3)', 3],
    ['ODD(2)', 3],
    ['ODD(-1)', -1],
    ['ODD(-2)', -3],
    ['CEILING(2.5,1)', 3],
    ['CEILING(-2.5,-2)', -4],
    ['CEILING(-2.5,2)', -2],
    ['CEILING(1.5,0.1)', 1.5],
    ['CEILING(0.234,0.01)', 0.24],
    ['CEILING.MATH(24.3,5)', 25],
    ['CEILING.MATH(6.7)', 7],
    ['CEILING.MATH(-8.1,2)', -8],
    ['CEILING.MATH(-5.5,2,-1)', -6],
    ['FLOOR(3.7,2)', 2],
    ['FLOOR(-2.5,-2)', -2],
    ['FLOOR(1.58,0.1)', 1.5],
    ['FLOOR(0.234,0.01)', 0.23],
    ['FLOOR.MATH(24.3,5)', 20],
    ['FLOOR.MATH(6.7)', 6],
    ['FLOOR.MATH(-8.1,2)', -10],
    ['FLOOR.MATH(-5.5,2,-1)', -4],
  ]);

  examples('integers', {}, [
    ['MOD(3,2)', 1],
    ['MOD(-3,2)', 1],
    ['MOD(3,-2)', -1],
    ['MOD(-3,-2)', -1],
    ['QUOTIENT(5,2)', 2],
    ['QUOTIENT(4.5,3.1)', 1],
    ['QUOTIENT(-10,3)', -3],
    ['FACT(5)', 120],
    ['FACT(1.9)', 1],
    ['FACT(0)', 1],
    ['FACT(1)', 1],
    ['FACTDOUBLE(6)', 48],
    ['FACTDOUBLE(7)', 105],
    ['COMBIN(8,2)', 28],
    ['GCD(5,2)', 1],
    ['GCD(24,36)', 12],
    ['GCD(7,1)', 1],
    ['GCD(5,0)', 5],
    ['LCM(5,2)', 10],
    ['LCM(24,36)', 72],
  ]);

  examples(
    'totals of ranges',
    {
      ...column('A2', [5, 15, 30]),
      ...column('C2', [3, 8, 1]),
      ...column('D2', [4, 6, 9]),
      ...column('F2', [2, 6, 5]),
      ...column('G2', [7, 7, 3]),
    },
    [
      ['PRODUCT(A2:A4)', 2250],
      ['PRODUCT(A2:A4,2)', 4500],
      ['PRODUCT(B2:B4)', 0],
      ['SUMPRODUCT(C2:D4,F2:G4)', 156],
      ['SUMPRODUCT(A2:A4)', 50],
      ['SUMPRODUCT(3,4)', 12],
    ],
  );

  examples(
    'SUMIF over amounts',
    {
      ...column('A2', [100000, 200000, 300000, 400000]),
      ...column('B2', [7000, 14000, 21000, 28000]),
      C2: 250000,
    },
    [
      ['SUMIF(A2:A5,">160000",B2:B5)', 63000],
      ['SUMIF(A2:A5,">160000")', 900000],
      ['SUMIF(A2:A5,300000,B2:B5)', 21000],
      ['SUMIF(A2:A5,">"&C2,B2:B5)', 49000],
      // A sum range of another size has the shape of the criteria range.
      ['SUMIF(A2:A5,">160000",B2)', 63000],
    ],
  );

  examples(
    'SUMIF over categories',
    {
      ...column('A2', ['Vegetables', 'Vegetables', 'Fruits', null]),
      ...column('A6', ['Vegetables', 'Fruits']),
      ...column('B2', ['Tomatoes', 'Celery', 'Oranges', 'Butter']),
      ...column('B6', ['Carrots', 'Apples']),
      ...column('C2', [2300, 5500, 800, 400, 4200, 1200]),
    },
    [
      ['SUMIF(A2:A7,"Fruits",C2:C7)', 2000],
      ['SUMIF(A2:A7,"Vegetables",C2:C7)', 12000],
      ['SUMIF(B2:B7,"*es",C2:C7)', 4300],
      ['SUMIF(A2:A7,"",C2:C7)', 400],
    ],
  );

  examples(
    'SUMIFS',
    {
      ...column('A2', [5, 4, 15, 3, 22, 12, 10, 33]),
      ...column('B2', [
        'Apples',
        'Apples',
        'Artichokes',
        'Artichokes',
        'Bananas',
        'Bananas',
        'Carrots',
        'Carrots',
      ]),
      ...column('C2', [
        'Tom',
        'Sarah',
        'Tom',
        'Sarah',
        'Tom',
        'Sarah',
        'Tom',
        'Sarah',
      ]),
    },
    [
      ['SUMIFS(A2:A9,B2:B9,"=A*",C2:C9,"Tom")', 20],
      ['SUMIFS(A2:A9,B2:B9,"<>Bananas",C2:C9,"Tom")', 30],
    ],
  );
});

describe('statistical functions', () => {
  const sample = [1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299];

  examples(
    'spread',
    column('A2', sample),
    [
      ['STDEV.S(A2:A11)', 27.46392],
      ['STDEV(A2:A11)', 27.46392],
      ['STDEV.P(A2:A11)', 26.05456],
      ['STDEVP(A2:A11)', 26.05456],
      ['VAR.S(A2:A11)', 754.2667],
      ['VAR(A2:A11)', 754.2667],
      ['VAR.P(A2:A11)', 678.84],
      ['VARP(A2:A11)', 678.84],
    ],
    true,
  );

  examples(
    'deviations and means',
    column('A2', [4, 5, 8, 7, 11, 4, 3]),
    [
      ['AVEDEV(4,5,6,7,5,4,3)', 1.020408],
      ['DEVSQ(4,5,8,7,11,4,3)', 48],
      ['GEOMEAN(A2:A8)', 5.476987],
      ['HARMEAN(A2:A8)', 5.028376],
    ],
    true,
  );

  examples(
    'averages and counts of all values',
    {
      ...column('A2', [10, 7, 9, 2, 'Not available']),
      ...column('B2', [
        'Sales',
        39790,
        null,
        19,
        22.24,
        true,
        error('#DIV/0!'),
      ]),
      ...column('C2', [0, 0.2, 0.5, 0.4, true]),
      ...column('D2', [false, 0.2, 0.5, 0.4, 0.8]),
      ...column('E2', ['', null, 1]),
    },
    [
      ['AVERAGEA(A2:A6)', 5.6],
      ['AVERAGEA(A2:A5,A7)', 7],
      ['COUNTA(B2:B8)', 6],
      ['COUNTA(1,"a",TRUE)', 3],
      ['COUNTBLANK(E2:E4)', 2],
      ['MAXA(C2:C6)', 1],
      ['MINA(D2:D6)', 0],
    ],
  );

  examples(
    'order statistics',
    {
      ...column('A2', [3, 4, 5, 2, 3]),
      ...column('B2', [4, 5, 6, 4, 7]),
      ...column('C2', [3, 4, 5, 2, 3, 4, 6, 4, 7]),
      ...column('D2', [1, 4, 8, 3, 7, 12, 54, 8, 23]),
      ...column('E2', [1, 2, 3, 4, 5, 6]),
      ...column('F2', [5.6, 4, 4, 3, 2, 4]),
      ...column('G2', [1, 3, 2, 4]),
      ...column('H2', [1, 2, 4, 7, 8, 9, 10, 12]),
    },
    [
      ['LARGE(A2:B6,3)', 5],
      ['LARGE(A2:B6,7)', 4],
      ['SMALL(C2:C10,4)', 4],
      ['SMALL(D2:D10,2)', 3],
      ['MEDIAN(E2:E6)', 3],
      ['MEDIAN(E2:E7)', 3.5],
      ['MODE.SNGL(F2:F7)', 4],
      ['MODE(F2:F7)', 4],
      ['MODE(3,1,1,3)', 3],
      ['PERCENTILE.INC(G2:G5,0.3)', 1.9],
      ['PERCENTILE(G2:G5,0.3)', 1.9],
      ['QUARTILE(H2:H9,1)', 3.5],
      ['QUARTILE.INC(H2:H9,1)', 3.5],
    ],
  );

  examples(
    'ranks and correlation',
    {
      ...column('A2', [7, 3.5, 3.5, 1, 2]),
      ...column('B2', [89, 88, 92, 101, 94, 97, 95]),
      ...column('C2', [3, 2, 4, 5, 6]),
      ...column('D2', [9, 7, 12, 15, 17]),
    },
    [
      ['RANK(A3,A2:A6,1)', 3],
      ['RANK(A2,A2:A6,1)', 5],
      ['RANK.EQ(A2,A2:A6)', 1],
      ['RANK.EQ(A3,A2:A6,1)', 3],
      ['RANK.AVG(94,B2:B8)', 4],
      ['RANK.AVG(A3,A2:A6,1)', 3.5],
      ['CORREL(C2:C6,D2:D6)', 0.997054486],
    ],
    true,
  );

  examples(
    'AVERAGEIF',
    {
      ...column('A2', [100000, 200000, 300000, 400000]),
      ...column('B2', [7000, 14000, 21000, 28000]),
      ...column('C2', [
        'East',
        'West',
        'North',
        'South (New Office)',
        'MidWest',
      ]),
      ...column('D2', [45678, 23789, -4789, 0, 9678]),
    },
    [
      ['AVERAGEIF(B2:B5,"<23000")', 14000],
      ['AVERAGEIF(A2:A5,"<250000")', 150000],
      ['AVERAGEIF(A2:A5,"<95000")', error('#DIV/0!')],
      ['AVERAGEIF(A2:A5,">250000",B2:B5)', 24500],
      ['AVERAGEIF(C2:C6,"=*West",D2:D6)', 16733.5],
      ['AVERAGEIF(C2:C6,"<>*(New Office)",D2:D6)', 18589],
    ],
  );

  examples(
    'AVERAGEIFS',
    {
      ...column('B2', [75, 94, 86, 'Incomplete']),
      ...column('C2', [85, 80, 93, 75]),
      ...column('D2', [87, 88, 'Incomplete', 75]),
    },
    [
      ['AVERAGEIFS(B2:B5,B2:B5,">70",B2:B5,"<90")', 80.5],
      ['AVERAGEIFS(C2:C5,C2:C5,">95")', error('#DIV/0!')],
      ['AVERAGEIFS(D2:D5,D2:D5,"<>Incomplete",D2:D5,">80")', 87.5],
    ],
  );

  examples(
    'COUNTIF',
    {
      ...column('A2', ['apples', 'oranges', 'peaches', 'apples']),
      ...column('B2', [32, 54, 75, 86]),
    },
    [
      ['COUNTIF(A2:A5,"apples")', 2],
      ['COUNTIF(A2:A5,A4)', 1],
      ['COUNTIF(A2:A5,A2)+COUNTIF(A2:A5,A3)', 3],
      ['COUNTIF(B2:B5,">55")', 2],
      ['COUNTIF(B2:B5,"<>"&B4)', 3],
      ['COUNTIF(B2:B5,">=32")-COUNTIF(B2:B5,">85")', 3],
      ['COUNTIF(A2:A5,"*")', 4],
      ['COUNTIF(A2:A5,"?????es")', 2],
    ],
  );

  examples(
    'COUNTIFS',
    {
      ...column('B2', ['Yes', 'Yes', 'Yes', 'No']),
      ...column('C2', ['No', 'Yes', 'Yes', 'Yes']),
      ...column('D2', ['No', 'No', 'Yes', 'Yes']),
    },
    [
      ['COUNTIFS(B2:D2,"=Yes")', 1],
      ['COUNTIFS(B2:B5,"=Yes",C2:C5,"=Yes")', 2],
      ['COUNTIFS(B5:D5,"=Yes",B3:D3,"=Yes")', 1],
    ],
  );

  examples(
    'MAXIFS and MINIFS',
    {
      ...column('A2', [89, 93, 96, 85, 91, 88]),
      ...column('B2', [1, 2, 2, 3, 1, 1]),
    },
    [
      ['MAXIFS(A2:A7,B2:B7,1)', 91],
      ['MINIFS(A2:A7,B2:B7,1)', 88],
      ['MAXIFS(A2:A7,B2:B7,9)', 0],
    ],
  );
});

describe('criteria of the conditional totals', () => {
  const cells = {
    ...column('A1', [5, '5', 'Apple', 'apple pie', null, true, 'a*b', 10]),
    ...column('A9', [error('#N/A'), '']),
  };

  examples('matching', cells, [
    // A number matches numbers and text that reads as that number.
    ['COUNTIF(A1:A10,5)', 2],
    ['COUNTIF(A1:A10,"5")', 2],
    ['COUNTIF(A1:A10,"=5")', 2],
    // Comparisons of numbers ignore text, and of text ignore numbers.
    ['COUNTIF(A1:A10,">=5")', 2],
    ['COUNTIF(A1:A8,"<b")', 4],
    // Text matches without regard to case; `~` escapes a wildcard.
    ['COUNTIF(A1:A10,"APPLE")', 1],
    ['COUNTIF(A1:A10,"apple*")', 2],
    ['COUNTIF(A1:A10,"a~*b")', 1],
    ['COUNTIF(A1:A10,"a*b")', 1],
    ['COUNTIF(A1:A10,TRUE)', 1],
    ['COUNTIF(A1:A10,"TRUE")', 1],
    ['COUNTIF(A1:A10,#N/A)', 1],
    // Blank criteria: empty cells, with or without empty text.
    ['COUNTIF(A1:A10,"")', 2],
    ['COUNTIF(A1:A10,"=")', 1],
    ['COUNTIF(A1:A10,"<>")', 9],
    ['COUNTIF(A1:A10,"<>5")', 8],
    ['COUNTIF(A1:A12,"<>Apple")', 11],
    ['COUNTIF(B:B,"")', 100000],
    ['SUMIF(A1:A8,"<>5",A1:A8)', 10],
  ]);
});

describe('logical functions', () => {
  examples('AND, OR, NOT and XOR', { A2: 50, A3: 'text', A4: true }, [
    ['AND(TRUE,TRUE)', true],
    ['AND(TRUE,FALSE)', false],
    ['AND(2+2=4,2+3=5)', true],
    ['AND(A2>1,A2<100)', true],
    ['AND(A2:A4)', true],
    ['OR(TRUE)', true],
    ['OR(1+1=1,2+2=5)', false],
    ['OR(TRUE,FALSE,TRUE)', true],
    ['OR(0,"false")', false],
    ['NOT(FALSE)', true],
    ['NOT(1+1=2)', false],
    ['XOR(3>0,2<9)', false],
    ['XOR(3>12,4>6)', false],
    ['XOR(TRUE,FALSE,FALSE)', true],
    ['TRUE()', true],
    ['FALSE()', false],
  ]);

  examples('IFS and SWITCH', { A2: 93, A3: 89, A4: 50, B2: 7 }, [
    ['IFS(A2>89,"A",A2>79,"B",A2>69,"C",A2>59,"D",TRUE,"F")', 'A'],
    ['IFS(A3>89,"A",A3>79,"B",A3>69,"C",A3>59,"D",TRUE,"F")', 'B'],
    ['IFS(A4>89,"A",A4>79,"B",A4>69,"C",A4>59,"D",TRUE,"F")', 'F'],
    ['IFS(A4>89,"A")', error('#N/A')],
    ['IFS(FALSE,1/0,TRUE,2)', 2],
    ['SWITCH(2,1,"Sunday",2,"Monday",3,"Tuesday","No match")', 'Monday'],
    ['SWITCH(99,1,"Sunday",2,"Monday",3,"Tuesday","No match")', 'No match'],
    ['SWITCH(99,1,"Sunday",2,"Monday",3,"Tuesday")', error('#N/A')],
    ['SWITCH(B2,7,"Weekend",1,"Weekend","Weekday")', 'Weekend'],
    ['SWITCH("b","A",1,"B",2)', 2],
  ]);

  examples(
    'IFERROR and IFNA',
    { ...column('A2', [210, 55, null]), ...column('B2', [35, 0, 23]) },
    [
      ['IFERROR(A2/B2,"Error in calculation")', 6],
      ['IFERROR(A3/B3,"Error in calculation")', 'Error in calculation'],
      ['IFERROR(A4/B4,"Error in calculation")', 0],
      ['IFNA(#N/A,"Not found")', 'Not found'],
      ['IFNA(1/0,"Not found")', error('#DIV/0!')],
      ['IFNA(A2,"Not found")', 210],
    ],
  );
});

describe('function errors', () => {
  const cells = {
    ...column('A1', [1, 2, 3]),
    ...column('B1', [4, 5]),
    ...column('D1', [7, 7]),
  };

  it.each<[string, string]>([
    // Text where a number is expected.
    ['ABS("a")', '#VALUE!'],
    ['ROUND("x",1)', '#VALUE!'],
    ['MEDIAN("a")', '#VALUE!'],
    ['AND("a")', '#VALUE!'],
    ['AND(C1:C3)', '#VALUE!'],
    ['NOT("a")', '#VALUE!'],
    // Ranges of other shapes.
    ['SUMPRODUCT(A1:A3,B1:B2)', '#VALUE!'],
    ['SUMIFS(A1:A3,B1:B2,">0")', '#VALUE!'],
    ['COUNTIFS(A1:A3,">0",B1:B2,">0")', '#VALUE!'],
    ['CORREL(A1:A3,B1:B2)', '#N/A'],
    ['SUMIF(1,">0")', '#VALUE!'],
    // Numbers outside the domain of a function.
    ['SQRT(-16)', '#NUM!'],
    ['LN(0)', '#NUM!'],
    ['LOG(-1)', '#NUM!'],
    ['ACOS(2)', '#NUM!'],
    ['ACOSH(0.5)', '#NUM!'],
    ['ATANH(1)', '#NUM!'],
    ['FACT(-1)', '#NUM!'],
    ['FACT(171)', '#NUM!'],
    ['COMBIN(2,3)', '#NUM!'],
    ['GCD(-1,2)', '#NUM!'],
    ['FLOOR(2.5,-2)', '#NUM!'],
    ['CEILING(2.5,-2)', '#NUM!'],
    ['MROUND(5,-2)', '#NUM!'],
    ['LARGE(A1:A3,4)', '#NUM!'],
    ['SMALL(A1:A3,0)', '#NUM!'],
    ['MEDIAN(C1:C3)', '#NUM!'],
    ['PERCENTILE(A1:A3,1.5)', '#NUM!'],
    ['QUARTILE(A1:A3,5)', '#NUM!'],
    ['GEOMEAN(0,1)', '#NUM!'],
    ['POWER(0,0)', '#NUM!'],
    // Division by zero.
    ['MOD(1,0)', '#DIV/0!'],
    ['QUOTIENT(1,0)', '#DIV/0!'],
    ['LOG(10,1)', '#DIV/0!'],
    ['ATAN2(0,0)', '#DIV/0!'],
    ['FLOOR(1,0)', '#DIV/0!'],
    ['STDEV.S(1)', '#DIV/0!'],
    ['VAR(C1:C3)', '#DIV/0!'],
    ['CORREL(A1:A2,D1:D2)', '#DIV/0!'],
    // Nothing found.
    ['MODE(1,2,3)', '#N/A'],
    ['RANK(9,A1:A3)', '#N/A'],
    // Errors pass on.
    ['ABS(1/0)', '#DIV/0!'],
    ['ROUND(1,#N/A)', '#N/A'],
    ['SUMIF(A1:A3,">0",#N/A)', '#N/A'],
    ['IFERROR(#N/A,1/0)', '#DIV/0!'],
  ])('%s gives %s', (formula, code) => {
    expect(calculate(formula, cells)).toEqual(error(code));
  });

  it('passes on the errors of the cells a conditional total adds', () => {
    const cells = {
      ...column('A1', ['x', 'y']),
      ...column('B1', [1, error('#N/A')]),
    };
    expect(calculate('SUMIF(A1:A2,"x",B1:B2)', cells)).toBe(1);
    expect(calculate('SUMIF(A1:A2,"y",B1:B2)', cells)).toEqual(error('#N/A'));
  });

  it.each([
    'ABS()',
    'ABS(1,2)',
    'PI(1)',
    'ROUND(1)',
    'IFS(TRUE)',
    'IFS(TRUE,1,FALSE)',
    'SWITCH(1,2)',
    'SUMIF(A1:A2)',
    'SUMIFS(A1:A2,B1:B2)',
    'COUNTIFS(A1:A2,1,B1:B2)',
    'NOT()',
  ])('refuses %s, with the wrong number of arguments', (formula) => {
    expect(parse(formula)).toBeNull();
  });
});

describe('function names in files', () => {
  it('reads the prefix that files give newer functions', () => {
    expect(print(parse('_xlfn.STDEV.S(A1:A2)+_XLFN.ifs(TRUE,1)')!)).toBe(
      'STDEV.S(A1:A2)+IFS(TRUE,1)',
    );
  });

  it('writes the prefix for files, for newer functions only', () => {
    const ast = parse('IFS(A1>0,STDEV.S(A1:A2))+SUM(A1)+MAXIFS(A1,B1,1)')!;
    expect(print(ast, { file: true })).toBe(
      '_xlfn.IFS(A1>0,_xlfn.STDEV.S(A1:A2))+SUM(A1)+_xlfn.MAXIFS(A1,B1,1)',
    );
  });
});
