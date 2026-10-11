import { createEngine } from '../public/js/formula.js';
const cells = {
  '0:0': {v:'Label'}, '1:0': {v:10}, '2:0': {v:20}, '3:0': {v:30},
  '1:1': {f:'=SUM(A2:A4)'}, '1:2': {f:'=AVERAGE(A2:A4)'}, '1:3': {f:'=MAX(A2:A4)-MIN(A2:A4)'},
  '2:1': {f:'=A2*2+5'}, '2:2': {f:'=IF(A3>15,"big","small")'}, '2:3': {f:'=ROUND(10/3,2)'},
  '3:1': {f:'=COUNTIF(A2:A4,">15")'}, '3:2': {f:'=SUMIF(A2:A4,">15")'}, '3:3': {f:'=A2&"/"&A3'},
  '4:1': {f:'=SUM(B2:B3)'}, '4:2': {f:'=1/0'}, '4:3': {f:'=LEFT("hello",3)&UPPER("x")'},
  '5:1': {f:'=COUNTIF(A1:A1,"Lab*")'}, '5:3': {f:'=100%'},
  '6:1': {f:'=TODAY()'}, '6:3': {f:'=B2+B3'}, '7:3': {f:'=A2+1'}, '7:0': {v:0.5}, '7:1': {f:'=A8*100'},
  '8:1': {f:'=B8'}, '9:1': {f:'=(2+3)*4-6/2'}, '9:2': {f:'=IF(AND(true,1>0),"y","n")'}, '9:3': {f:'=VLOOKUP(20,A2:A4,1,false)'}
};
const e = createEngine(cells);
const t = (r,c,label) => console.log(String(label).padEnd(24), '→', JSON.stringify(e.resolve(r,c).display));
t(1,1,'SUM(A2:A4)'); t(1,2,'AVERAGE'); t(1,3,'MAX-MIN'); t(2,1,'A2*2+5'); t(2,2,'IF');
t(2,3,'ROUND(10/3,2)'); t(3,1,'COUNTIF >15'); t(3,2,'SUMIF >15'); t(3,3,'concat &');
t(4,1,'SUM of formulas'); t(4,2,'1/0'); t(4,3,'LEFT+UPPER'); t(5,1,'wildcard Lab*');
t(5,3,'100%'); t(6,1,'TODAY()'); t(6,3,'B2+B3'); t(7,3,'A2+1'); t(7,1,'0.5*100');
t(8,1,'=B8 cycle'); t(9,1,'(2+3)*4-6/2'); t(9,2,'AND/IF'); t(9,3,'VLOOKUP');
