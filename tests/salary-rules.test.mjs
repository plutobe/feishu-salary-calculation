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
  calcDayOvertime,
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

// 弹性打卡加班推导：加班从 max(18:00, 上班卡+9.5h) 起算
assert.equal(Number(rules.calcDayOvertime('正常(08:50),正常(21:26)').toFixed(2)), 3.10); // >2h
assert.equal(Number(rules.calcDayOvertime('正常(08:39),正常(19:19)').toFixed(2)), 1.17); // <2h
assert.equal(Number(rules.calcDayOvertime('正常(07:47),正常(22:58)').toFixed(2)), 4.97); // 早到按18:00起算
assert.equal(Number(rules.calcDayOvertime('休息打卡(09:10,17:36);加班(09:30-12:00,13:30-17:29)').toFixed(2)), 6.48); // 休息打卡按已审批时段求和
assert.equal(Number(rules.calcDayOvertime('正常(08:40),正常(18:12)').toFixed(2)), 0.03); // 弹性下仅2分钟

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

// 公积金个人扣款：未设公积金基数/社保基数时按正式月薪×个人12%
assert.equal(Number(firstDayHire.housingFundDeduction.toFixed(2)), 1440.00);

// 餐补：2天加班>2h，每天15元
const mealEmp = rules.calcEmployee({
  ...baseEmployee,
  monthlySalary: 12000,
  daily: [
    { date: '2026-06-01', overtimeHours: 3.1 }, // >2h
    { date: '2026-06-02', overtimeHours: 1.17 },// <2h
    { date: '2026-06-03', overtimeHours: 0 },
    { date: '2026-06-04', overtimeHours: 2.5 },  // >2h
  ],
});
assert.equal(mealEmp.mealDays, 2);
assert.equal(mealEmp.mealAllowance, 30);

// 公积金基数跟随社保基数：设 socialInsuranceBase=10000，未设公积金基数 → 按10000×12%
const withSiBase = rules.calcEmployee({ ...baseEmployee, socialInsuranceBase: 10000 });
assert.equal(Number(withSiBase.housingFundDeduction.toFixed(2)), 1200.00);
// 设公积金基数则优先用公积金基数
const withHfBase = rules.calcEmployee({ ...baseEmployee, socialInsuranceBase: 10000, housingFundBase: 8000 });
assert.equal(Number(withHfBase.housingFundDeduction.toFixed(2)), 960.00);
// payHousingFund=false 不扣公积金
assert.equal(rules.calcEmployee({ ...baseEmployee, payHousingFund: false }).housingFundDeduction, 0);

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
