import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('../js/script.js', import.meta.url), 'utf8');
const calculationSource = source.split('// ======= Render =======')[0] + `
globalThis.salaryRules = {
  extractLeaveHoursFromDaily,
  countPayableDays,
  calcPaidAttendanceDays,
  calcDeductibleAbsenceHours,
  calcEmployee,
};`;
const context = { console };
vm.runInNewContext(calculationSource, context);

const rules = context.salaryRules;
const juneDates = Array.from({ length: 30 }, (_, index) => ({
  date: `2026-06-${String(index + 1).padStart(2, '0')}`,
}));

assert.equal(rules.extractLeaveHoursFromDaily('调休假(13:30-18:00)', '调休假'), 4.5);
assert.equal(rules.extractLeaveHoursFromDaily('调休假(08:30-12:00,13:30-18:00)', '调休假'), 8);
assert.equal(rules.countPayableDays('2026-06-15', '2026-06-30'), 12);
assert.equal(rules.countPayableDays('2026-05-02', '2026-05-03'), 1);

assert.equal(rules.calcPaidAttendanceDays({
  requiredDays: 21,
  requiredHours: 168,
  actualDays: 20.625,
  compHours: 3,
}), 21);

assert.equal(rules.calcDeductibleAbsenceHours({ sickHours: 0, personalHours: 1 }), 1);

const baseEmployee = {
  monthlySalary: 12000,
  probationSalary: 9600,
  probationMonths: 3,
  daily: juneDates,
  sickHours: 0,
  compHours: 0,
  personalHours: 0,
  paySocialInsurance: false,
  requiredDays: 21,
  requiredHours: 168,
  actualDays: 21,
};

const firstDayHire = rules.calcEmployee({ ...baseEmployee, hireDate: '2026-06-01' });
assert.equal(firstDayHire.isFirstDayHire, true);
assert.equal(firstDayHire.base, 9600);

const liAo = rules.calcEmployee({
  ...baseEmployee,
  hireDate: '2026-06-15',
  actualDays: 9,
  personalHours: 16,
});
assert.equal(liAo.payableDays, 12);
assert.equal(Number(liAo.base.toFixed(2)), 5296.55);
assert.equal(Number(liAo.personalDeduction.toFixed(2)), 882.76);

console.log('salary rule tests passed');
