import { runQuery } from '../public/js/query.js';

const db = {
  tables: [
    { name: 'Departments', fields: [], rows: [{ Id: 1, Name: 'Engineering' }, { Id: 2, Name: 'Design' }] },
    { name: 'People', fields: [], rows: [{ Name: 'Ben', Role: 'Eng', Dept: 1 }, { Name: 'Amelia', Role: 'Design', Dept: 2 }] }
  ]
};

let fails = 0;
const check = (n, c) => { console.log((c ? '  ok   ' : '  FAIL ') + n); if (!c) fails++; };

const all = runQuery(db, 'SELECT * FROM People');
check('select all', all.rows.length === 2 && all.columns.includes('People.Name'));

const filtered = runQuery(db, "SELECT Name, Role FROM People WHERE Role = 'Eng'");
check('where equals', filtered.rows.length === 1 && filtered.rows[0].Role === 'Eng');

const joined = runQuery(db, 'SELECT People.Name, Departments.Name FROM People JOIN Departments ON People.Dept = Departments.Id');
check('join', joined.rows.length === 2 && joined.rows.some(r => r['Departments.Name'] === 'Engineering'));

const bad = runQuery(db, 'DROP TABLE People');
check('rejects unknown sql', !!bad.error);

console.log(fails ? fails + ' FAILED' : 'QUERY CHECKS PASSED');
process.exit(fails ? 1 : 0);
