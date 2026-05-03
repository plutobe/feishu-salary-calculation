// ======= State =======
let employees = [];

let socialInsurance = {
  pension: { unit: 16, personal: 8 },
  medical: { unit: 9.5, personal: 2 },
  maternity: { unit: 0.5, personal: 0 },
  unemployment: { unit: 0.5, personal: 0.5 },
  injury: { unit: 0.6, personal: 0 },
  medicalExtra: 5
};

let probationSettings = {
  discount: 80, // 试用期折扣百分比
  months: 3     // 默认试用期月数
};

// ======= Parse Daily Attendance =======
function parseDailyStatus(val) {
  if (!val || val === '-') return { type: 'none', title: '-' };
  const s = String(val);

  if (s.includes('入职日')) return { type: 'entry', title: '入职日' };
  if (s.includes('休息')) return { type: 'rest', title: '休息' };
  if (s.includes('调休假')) return { type: 'comp', title: '调休假' };

  // 判断是否有标记过缺勤
  if (s.includes('缺卡')) return { type: 'miss', title: '缺卡' };
  if (s.includes('旷工')) return { type: 'absent', title: '旷工' };
  if (s.includes('外出')) return { type: 'normal', title: '外出' };

  const hasOT = s.includes('加班');
  if (s.includes('正常')) {
    return { type: hasOT ? 'overtime' : 'normal', title: hasOT ? '加班' : '正常' };
  }

  return { type: 'none', title: s.substring(0, 10) };
}

/** Parse 调休假 hours from daily value like "请假(-),请假(-);调休假(08:30-12:00,13:30-18:00)" */
function extractCompHoursFromDaily(val) {
  if (!val) return 0;
  const s = String(val);
  const match = s.match(/调休假\(([^)]+)\)/);
  if (!match) return 0;
  const ranges = match[1].split(',');
  let total = 0;
  for (const range of ranges) {
    const parts = range.split('-');
    if (parts.length === 2) {
      const [start, end] = parts;
      const startMin = timeToMinutes(start);
      const endMin = timeToMinutes(end);
      if (endMin > startMin) total += (endMin - startMin) / 60;
    }
  }
  return total;
}

function timeToMinutes(t) {
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}

function parseExcel(data) {
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

  // Row 1: category headers, Row 2: column headers, Row 3+: data
  const dataRows = rows.slice(2).filter(r => r[0] && String(r[0]).trim());

  const dailyColStart = 46; // 0-indexed, column AU = 46
  const dailyColEnd = 75;   // 0-indexed, column BX = 75

  // Parse daily dates from header row (row 1, index 0)
  const headerRow = rows[1];
  const dailyDates = [];
  for (let d = dailyColStart; d <= dailyColEnd; d++) {
    const header = String(headerRow[d] || '');
    // Extract date from "2026-04-01 星期三" format
    const dateMatch = header.match(/(\d{4}-\d{2}-\d{2})/);
    dailyDates.push(dateMatch ? dateMatch[1] : null);
  }

  return dataRows.map((r, idx) => {
    const absenceHours = parseFloat(r[30]) || 0; // Col AF (31), 0-indexed 30
    const monthlyDays = parseFloat(r[13]) || 0;   // Col N (14), 0-indexed 13
    const monthlyHours = parseFloat(r[14]) || 0;  // Col O (15), 0-indexed 14

    // Parse hire date from Col 10
    const hireDateRaw = r[10];
    let hireDate = null;
    if (hireDateRaw) {
      if (hireDateRaw instanceof Date) {
        hireDate = hireDateRaw.toISOString().split('T')[0];
      } else {
        const dateStr = String(hireDateRaw).trim();
        const match = dateStr.match(/(\d{4}-\d{2}-\d{2})/);
        if (match) hireDate = match[1];
      }
    }

    // Build daily attendance with dates
    const daily = [];
    let autoCompHours = 0;
    for (let d = dailyColStart; d <= dailyColEnd; d++) {
      const val = r[d];
      const status = parseDailyStatus(val);
      status.date = dailyDates[d - dailyColStart]; // Add date to each day
      daily.push(status);
      if (status.type === 'comp') {
        autoCompHours += extractCompHoursFromDaily(val);
      }
    }

    // Pre-fill: round comp hours to half-day unit
    const halfDay = monthlyHours && monthlyDays ? (monthlyHours / monthlyDays / 2) : 4;
    const compHours = autoCompHours > 0 ? Math.round(autoCompHours / halfDay) * halfDay : 0;

    return {
      id: idx,
      name: String(r[0] || '').trim(),
      employeeId: String(r[1] || '').trim(),
      dept: String(r[2] || '').trim(),
      hireDate: hireDate,
      monthlySalary: 0,
      probationMonths: 3, // default probation months
      paySocialInsurance: true,
      socialInsuranceBase: 0,
      requiredDays: monthlyDays,
      requiredHours: monthlyHours,
      actualDays: parseFloat(r[15]) || 0,   // Col P (16), 0-indexed 15
      actualHours: parseFloat(r[18]) || 0,  // Col S (19), 0-indexed 18
      absenceHours: absenceHours,
      overtimeHours: parseFloat(r[34]) || 0, // Col AJ (35), 0-indexed 34
      // editable
      sickHours: 0,
      compHours: compHours,
      personalHours: Math.max(0, absenceHours - compHours),
      daily: daily,
    };
  });
}

// ======= Half-day rounding =======
function calcHalfDayHours(emp) {
  if (!emp.requiredHours || !emp.requiredDays) return 4;
  return (emp.requiredHours / emp.requiredDays) / 2;
}

function roundToHalfDay(hours, emp) {
  const halfDay = calcHalfDayHours(emp);
  if (halfDay <= 0) return hours;
  return Math.round(hours / halfDay) * halfDay;
}

// ======= Calculations (按国家规定21.75天) =======
function calcEmployee(emp) {
  const formalSalary = emp.monthlySalary;
  const months = emp.probationMonths != null ? emp.probationMonths : probationSettings.months;
  const probationSalary = months === 0 ? formalSalary : (emp.probationSalary || (formalSalary * probationSettings.discount / 100));

  // Calculate probation end date
  let probationEndDate = null;
  if (emp.hireDate) {
    const hire = new Date(emp.hireDate);
    probationEndDate = new Date(hire);
    probationEndDate.setMonth(probationEndDate.getMonth() + months);
    probationEndDate.setDate(probationEndDate.getDate() - 1); // End day is day before
  }

  // Determine salary for each day and calculate weighted average
  let probationDays = 0;
  let formalDays = 0;
  const dailySalaries = []; // Store daily salary for each day

  emp.daily.forEach(d => {
    if (!d.date) {
      // No date info, use formal salary
      dailySalaries.push(formalSalary);
      formalDays++;
      return;
    }

    const dayDate = new Date(d.date);
    if (probationEndDate && dayDate <= probationEndDate) {
      // Probation period
      dailySalaries.push(probationSalary);
      probationDays++;
    } else {
      // Formal period
      dailySalaries.push(formalSalary);
      formalDays++;
    }
  });

  const totalDays = probationDays + formalDays;
  const weightedMonthlySalary = totalDays > 0
    ? (probationDays * probationSalary + formalDays * formalSalary) / totalDays
    : formalSalary;

  // 按实际出勤天数比例折算月薪
  // 整月员工 & 满勤 → 1.0；其他情况 → actual/21.75
  const firstDayDate = emp.daily && emp.daily[0] && emp.daily[0].date;
  const isFullMonthEmployee = !(emp.hireDate && firstDayDate
    && emp.hireDate.substring(0, 7) === firstDayDate.substring(0, 7));
  const isFullAttendance = emp.actualDays >= emp.requiredDays;
  const attendanceRatio = emp.actualDays > 0 && emp.requiredDays > 0
    ? (isFullMonthEmployee && isFullAttendance ? 1 : emp.actualDays / 21.75)
    : 1;
  const proratedBase = weightedMonthlySalary * attendanceRatio;

  const dailySalary = weightedMonthlySalary / 21.75;
  const hourlySalary = weightedMonthlySalary / 174; // 21.75 × 8

  // Sick leave deduction - use the daily salary for the day of leave
  let sickDeduction = 0;
  if (emp.sickHours > 0) {
    // Find which days are sick leave and use appropriate salary
    const sickHourlySalary = hourlySalary; // Use weighted average for simplicity
    sickDeduction = emp.sickHours * sickHourlySalary * 0.3;
  }

  // Personal leave deduction
  let personalDeduction = 0;
  if (emp.personalHours > 0) {
    const personalHourlySalary = hourlySalary;
    personalDeduction = emp.personalHours * personalHourlySalary;
  }

  // 社保个人扣款
  let siTotal = 0;
  let pensionDeduction = 0, medicalDeduction = 0, unemploymentDeduction = 0, maternityDeduction = 0, injuryDeduction = 0;
  if (emp.paySocialInsurance !== false) {
    const si = socialInsurance;
    const siBase = emp.socialInsuranceBase > 0 ? emp.socialInsuranceBase : formalSalary;
    pensionDeduction = siBase * si.pension.personal / 100;
    medicalDeduction = siBase * si.medical.personal / 100 + si.medicalExtra;
    unemploymentDeduction = siBase * si.unemployment.personal / 100;
    maternityDeduction = siBase * si.maternity.personal / 100;
    injuryDeduction = siBase * si.injury.personal / 100;
    siTotal = pensionDeduction + medicalDeduction + unemploymentDeduction + maternityDeduction + injuryDeduction;
  }

  // Net salary: prorate base salary by attendance ratio, then subtract deductions
  const netSalary = proratedBase - sickDeduction - personalDeduction - siTotal;

  return {
    dailySalary, hourlySalary, weightedMonthlySalary, attendanceRatio, proratedBase, probationEndDate,
    probationDays, formalDays, probationSalary,
    sickDeduction, personalDeduction,
    pensionDeduction, medicalDeduction, unemploymentDeduction, maternityDeduction, injuryDeduction,
    siTotal, netSalary
  };
}

// ======= Render =======
function render() {
  const container = document.getElementById('tableContainer');
  const emptyState = document.getElementById('emptyState');
  const summaryWrap = document.getElementById('summaryWrap');
  const statusInfo = document.getElementById('statusInfo');
  const exportBtn = document.getElementById('exportBtn');
  const settingsBtn = document.getElementById('settingsBtn');
  const clearBtn = document.getElementById('clearBtn');
  const employeeCount = document.getElementById('employeeCount');

  if (!employees.length) {
    container.style.display = 'none';
    emptyState.style.display = 'block';
    summaryWrap.style.display = 'none';
    statusInfo.textContent = '请上传 Excel 文件开始计算';
    employeeCount.style.display = 'none';
    exportBtn.style.display = 'none';
    settingsBtn.style.display = 'none';
    clearBtn.style.display = 'none';
    return;
  }

  container.style.display = '';
  emptyState.style.display = 'none';
  summaryWrap.style.display = '';
  exportBtn.style.display = '';
  settingsBtn.style.display = '';
  clearBtn.style.display = '';
  statusInfo.textContent = '';
  employeeCount.textContent = `共 ${employees.length} 名员工`;
  employeeCount.style.display = '';

  // Build table
  let html = '<table><thead><tr>' +
    '<th class="th-name">序号</th>' +
    '<th>姓名</th>' +
    '<th>部门</th>' +
    '<th>正式月薪</th>' +
    '<th>试用期<br><small>月薪/月数</small></th>' +
    '<th>缴纳社保</th>' +
    '<th>社保基数</th>' +
    '<th>入职日期</th>' +
    '<th>转正日期</th>' +
    '<th>应出勤<br><small>天</small></th>' +
    '<th>实际出勤<br><small>天</small></th>' +
    '<th>缺勤<br><small>小时</small></th>' +
    '<th>病假<br><small>小时</small></th>' +
    '<th>调休<br><small>小时</small></th>' +
    '<th>事假<br><small>小时</small></th>' +
    '<th>病假扣款</th>' +
    '<th>事假扣款</th>' +
    '<th>社保合计</th>' +
    '<th>实发工资</th>' +
    '<th>每日考勤</th>' +
    '<th>计算明细</th>' +
    '</tr></thead><tbody>';

  let totalSalary = 0;
  let totalSickDed = 0;
  let totalPersonalDed = 0;
  let totalSi = 0;
  let totalCompanySi = 0;
  let totalNet = 0;
  let totalEmployees = employees.length;

  employees.forEach((emp, idx) => {
    const calc = calcEmployee(emp);
    totalSalary += emp.monthlySalary;
    totalSickDed += calc.sickDeduction;
    totalPersonalDed += calc.personalDeduction;
    totalSi += calc.siTotal;
    totalNet += calc.netSalary;

    // Calculate company contribution
    const si = socialInsurance;
    const companySi = emp.monthlySalary * (si.pension.unit + si.medical.unit + si.unemployment.unit + si.maternity.unit + si.injury.unit) / 100;
    totalCompanySi += companySi;

    // Daily calendar — 7-column grid: 日 一 二 三 四 五 六
    const weekdayLabels = ['日', '一', '二', '三', '四', '五', '六'];
    const firstWeekday = 3; // 2026-04-01 is Wednesday (weekday 3, Sun=0)
    let dailyHtml = '<div class="daily-grid">';
    // Weekday header row
    for (let w = 0; w < 7; w++) {
      dailyHtml += `<div class="day-cell wd">${weekdayLabels[w]}</div>`;
    }
    // Empty cells before April 1 (Wednesday = column index 3)
    for (let w = 0; w < firstWeekday; w++) {
      dailyHtml += `<div style="visibility:hidden" class="day-cell none"></div>`;
    }
    // Calendar days
    for (let di = 0; di < emp.daily.length; di++) {
      const d = emp.daily[di];
      let cls = 'day-cell ';
      if (d.type === 'normal') cls += 'normal';
      else if (d.type === 'rest') cls += 'rest';
      else if (d.type === 'miss') cls += 'miss';
      else if (d.type === 'entry') cls += 'entry';
      else if (d.type === 'overtime') cls += 'overtime';
      else if (d.type === 'absent') cls += 'absent';
      else if (d.type === 'comp') cls += 'comp';
      else cls += 'none';
      dailyHtml += `<div class="${cls}" title="${di+1}日: ${d.title}">${di+1}</div>`;
    }
    dailyHtml += '</div>';

    const halfDay = calcHalfDayHours(emp);
    const probationMonths = emp.probationMonths != null ? emp.probationMonths : probationSettings.months;
    const probationSalary = probationMonths === 0 ? emp.monthlySalary : (emp.probationSalary || (emp.monthlySalary * probationSettings.discount / 100));
    const probationEndDateStr = calc.probationEndDate ? calc.probationEndDate.toISOString().split('T')[0] : '-';
    const hireDateStr = emp.hireDate || '-';
    const isProbation = calc.probationDays > 0 && calc.formalDays === 0;
    const isFormal = calc.probationDays === 0;
    const isTransition = calc.probationDays > 0 && calc.formalDays > 0;

    const probationDisabled = isFormal ? 'disabled' : '';
    html += `<tr>
      <td class="stat-value" style="text-align:center;color:var(--text-light)">${idx + 1}</td>
      <td class="td-name"><div class="employee-name">${escHtml(emp.name)}</div><div class="employee-dept">${escHtml(emp.employeeId)}</div></td>
      <td class="dept-col">${escHtml(emp.dept)}</td>
      <td><input type="number" class="salary-input" value="${emp.monthlySalary}" min="0" step="100" data-idx="${idx}" onchange="updateSalary(${idx}, this.value)"></td>
      <td>
        <input type="number" class="salary-input" value="${probationSalary}" min="0" step="100" style="width:80px${isFormal ? ';opacity:0.5' : ''}" ${probationDisabled} onchange="updateEmployeeProbationSalary(${idx}, this.value)">
        <br><small style="color:var(--text-light)">
          <input type="number" class="rate-input" value="${probationMonths}" min="0" max="12" step="1" style="width:40px${isFormal ? ';opacity:0.5' : ''}" ${probationDisabled} onchange="updateEmployeeProbationMonths(${idx}, this.value)">个月
        </small>
      </td>
      <td class="stat-value" style="text-align:center">
        <input type="checkbox" ${emp.paySocialInsurance !== false ? 'checked' : ''} onchange="updatePaySocialInsurance(${idx}, this.checked)">
      </td>
      <td>
        <input type="number" class="salary-input" value="${emp.socialInsuranceBase > 0 ? emp.socialInsuranceBase : emp.monthlySalary}" min="0" step="100" style="width:90px" onchange="updateSocialInsuranceBase(${idx}, this.value)">
      </td>
      <td class="stat-value" style="font-size:12px">${hireDateStr}</td>
      <td class="stat-value" style="font-size:12px">
        ${probationEndDateStr}
        ${isTransition ? '<br><small style="color:var(--warning)">本月转正</small>' : ''}
        ${isProbation ? '<br><small style="color:var(--primary)">试用期</small>' : ''}
        ${isFormal && emp.hireDate ? '<br><small style="color:var(--success)">已转正</small>' : ''}
      </td>
      <td class="stat-value">${emp.requiredDays}</td>
      <td class="stat-value">${emp.actualDays}</td>
      <td class="stat-value ${emp.absenceHours - emp.compHours > 0 ? 'stat-negative' : ''}">${(emp.absenceHours - emp.compHours).toFixed(1)}</td>
      <td>
        <input type="number" class="leave-input" value="${emp.sickHours}" min="0" max="${emp.absenceHours - emp.compHours}" step="${halfDay}" data-idx="${idx}" onchange="updateSick(${idx}, this.value)">
        <br><small style="color:var(--text-light)">0.5天=${halfDay}h</small>
      </td>
      <td>
        <input type="number" class="leave-input" value="${emp.compHours}" min="0" max="${emp.absenceHours}" step="${halfDay}" data-idx="${idx}" onchange="updateComp(${idx}, this.value)">
        <br><small style="color:var(--text-light)">0.5天=${halfDay}h</small>
      </td>
      <td>
        <span class="stat-value" id="personalHours_${idx}">${emp.personalHours.toFixed(1)}</span>
      </td>
      <td class="stat-negative stat-money">${fmtMoney(calc.sickDeduction)}</td>
      <td class="stat-negative stat-money">${fmtMoney(calc.personalDeduction)}</td>
      <td class="stat-negative stat-money" style="font-weight:700">${fmtMoney(calc.siTotal)}</td>
      <td class="stat-final">${fmtMoney(calc.netSalary)}</td>
      <td>${dailyHtml}</td>
      <td style="font-size:11px;line-height:1.5;white-space:normal;min-width:200px;color:var(--text-light)">
        ${isTransition ? `<div style="color:var(--warning)">本月转正：${calc.probationDays}天试用 + ${calc.formalDays}天正式</div>` : ''}
        <div>加权月薪 = ${fmtMoney(calc.weightedMonthlySalary)}</div>
        <div>日薪 = ${fmtMoney(calc.weightedMonthlySalary)} ÷ 21.75天 = <strong>${fmtMoney(calc.dailySalary)}</strong></div>
        <div>时薪 = ${fmtMoney(calc.weightedMonthlySalary)} ÷ 174h = <strong>${fmtMoney(calc.hourlySalary)}</strong></div>
        ${calc.attendanceRatio < 1 ? `<div>出勤比 = ${emp.actualDays}/21.75天 = ${(calc.attendanceRatio * 100).toFixed(1)}%</div>
        <div>折算月薪 = ${fmtMoney(calc.weightedMonthlySalary)} × ${(calc.attendanceRatio * 100).toFixed(1)}% = <strong>${fmtMoney(calc.proratedBase)}</strong></div>` : ''}
        ${emp.sickHours > 0 ? `<div>病假扣 = ${emp.sickHours}h × ${fmtMoney(calc.hourlySalary)} × 30% = <strong>${fmtMoney(calc.sickDeduction)}</strong></div>` : ''}
        ${emp.personalHours > 0 ? `<div>事假扣 = ${emp.personalHours}h × ${fmtMoney(calc.hourlySalary)} = <strong>${fmtMoney(calc.personalDeduction)}</strong></div>` : ''}
        ${emp.compHours > 0 ? '<div>调休假 = 不扣款</div>' : ''}
        <div>社保扣 = <strong>${emp.paySocialInsurance !== false ? fmtMoney(calc.siTotal) : '¥0.00'}</strong> <small>(${emp.paySocialInsurance !== false ? (emp.socialInsuranceBase > 0 ? '基数' + fmtMoney(emp.socialInsuranceBase) : '按正式薪资') : '不缴纳'})</small></div>
        <div style="color:var(--primary);font-weight:600;margin-top:2px">实发 = ${fmtMoney(calc.proratedBase)} − ${fmtMoney(calc.sickDeduction)} − ${fmtMoney(calc.personalDeduction)} − ${fmtMoney(calc.siTotal)} = ${fmtMoney(calc.netSalary)}</div>
      </td>
    </tr>`;
  });

  html += '</tbody></table>';
  container.innerHTML = html;

  // Summary
  const avgSalary = totalEmployees > 0 ? totalSalary / totalEmployees : 0;
  const avgNet = totalEmployees > 0 ? totalNet / totalEmployees : 0;
  const totalCost = totalNet + totalCompanySi;
  document.getElementById('summaryGrid').innerHTML = `
    <div class="summary-card"><div class="label">员工总数</div><div class="value">${totalEmployees}</div></div>
    <div class="summary-card"><div class="label">应发工资总额</div><div class="value primary">${fmtMoney(totalSalary)}</div></div>
    <div class="summary-card"><div class="label">病假扣款总额</div><div class="value danger">${fmtMoney(totalSickDed)}</div></div>
    <div class="summary-card"><div class="label">事假扣款总额</div><div class="value danger">${fmtMoney(totalPersonalDed)}</div></div>
    <div class="summary-card"><div class="label">个人社保总额</div><div class="value danger">${fmtMoney(totalSi)}</div></div>
    <div class="summary-card"><div class="label">实发工资总额</div><div class="value success">${fmtMoney(totalNet)}</div></div>
    <div class="summary-card"><div class="label">单位社保总额</div><div class="value" style="color:#e67e22">${fmtMoney(totalCompanySi)}</div></div>
    <div class="summary-card"><div class="label">企业用人总成本</div><div class="value" style="color:#8e44ad;font-size:20px">${fmtMoney(totalCost)}</div></div>
    <div class="summary-card"><div class="label">平均月薪</div><div class="value">${fmtMoney(avgSalary)}</div></div>
  `;
}

// ======= Event Handlers =======
function updateSalary(idx, val) {
  if (idx >= 0 && idx < employees.length) {
    employees[idx].monthlySalary = parseFloat(val) || 0;
    const emp = employees[idx];
    // Auto-sync socialInsuranceBase if not manually set
    if (emp.socialInsuranceBase === 0 || emp.socialInsuranceBase === undefined) {
      emp.socialInsuranceBase = emp.monthlySalary;
    }
    updateDerivedHours(emp);
    saveEmployeeCache(emp);
    render();
  }
}

function updateSick(idx, val) {
  if (idx >= 0 && idx < employees.length) {
    const emp = employees[idx];
    const halfDay = calcHalfDayHours(emp);
    let hours = parseFloat(val) || 0;
    hours = roundToHalfDay(hours, emp);
    hours = Math.max(0, Math.min(hours, emp.absenceHours - emp.compHours));
    emp.sickHours = hours;
    updateDerivedHours(emp);
    render();
  }
}

function updateComp(idx, val) {
  if (idx >= 0 && idx < employees.length) {
    const emp = employees[idx];
    const halfDay = calcHalfDayHours(emp);
    let hours = parseFloat(val) || 0;
    hours = roundToHalfDay(hours, emp);
    hours = Math.max(0, Math.min(hours, emp.absenceHours - emp.sickHours));
    emp.compHours = hours;
    updateDerivedHours(emp);
    render();
  }
}

function updateDerivedHours(emp) {
  // 调休假不算缺勤，缺勤仅含病假+事假
  const used = emp.sickHours + emp.compHours;
  emp.personalHours = Math.max(0, emp.absenceHours - used);
}

// ======= Modal Settings =======
function openSettings() {
  document.getElementById('settingsModal').classList.add('active');
}

function clearAll() {
  if (!confirm('确定要清空当前列表数据吗？（已保存的员工薪资设置不会被清除）')) return;
  employees = [];
  localStorage.removeItem(EMPLOYEES_CACHE_KEY);
  localStorage.removeItem(FILENAME_CACHE_KEY);
  document.getElementById('fileName').textContent = '';
  document.getElementById('fileName').parentElement.classList.remove('has-file');
  document.getElementById('fileInput').value = '';
  render();
}

function closeSettings() {
  document.getElementById('settingsModal').classList.remove('active');
}

// Close modal on overlay click
document.getElementById('settingsModal').addEventListener('click', function(e) {
  if (e.target === this) closeSettings();
});

// Close modal on Escape key
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') closeSettings();
});

function updateRate(type, field, val) {
  const v = parseFloat(val) || 0;
  if (type === 'medicalExtra') {
    socialInsurance.medicalExtra = v;
  } else {
    socialInsurance[type][field] = v;
  }
  render();
}

function updateProbationDiscount(val) {
  probationSettings.discount = parseFloat(val) || 80;
  render();
}

function updateProbationMonths(val) {
  const parsed = parseInt(val);
  probationSettings.months = isNaN(parsed) ? 3 : parsed;
  render();
}

function updateEmployeeProbationSalary(idx, val) {
  if (idx >= 0 && idx < employees.length) {
    employees[idx].probationSalary = parseFloat(val) || 0;
    saveEmployeeCache(employees[idx]);
    render();
  }
}

function updateEmployeeProbationMonths(idx, val) {
  if (idx >= 0 && idx < employees.length) {
    const parsed = parseInt(val);
    employees[idx].probationMonths = isNaN(parsed) ? probationSettings.months : parsed;
    saveEmployeeCache(employees[idx]);
    render();
  }
}

function updatePaySocialInsurance(idx, checked) {
  if (idx >= 0 && idx < employees.length) {
    employees[idx].paySocialInsurance = checked;
    saveEmployeeCache(employees[idx]);
    render();
  }
}

function updateSocialInsuranceBase(idx, val) {
  if (idx >= 0 && idx < employees.length) {
    employees[idx].socialInsuranceBase = parseFloat(val) || 0;
    saveEmployeeCache(employees[idx]);
    render();
  }
}

// ======= Cache =======
const CACHE_KEY = 'salaryCalc Employees';
const EMPLOYEES_CACHE_KEY = 'salaryCalc AllEmployees';
const FILENAME_CACHE_KEY = 'salaryCalc FileName';

function getCacheKey(emp) {
  return `${emp.name}|${emp.employeeId}`;
}

function loadEmployeeCache(employees) {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return;
    const cache = JSON.parse(raw);
    employees.forEach(emp => {
      const key = getCacheKey(emp);
      const saved = cache[key];
      if (saved) {
        if (saved.monthlySalary != null) emp.monthlySalary = saved.monthlySalary;
        if (saved.probationSalary != null) emp.probationSalary = saved.probationSalary;
        if (saved.probationMonths != null) emp.probationMonths = saved.probationMonths;
        if (saved.paySocialInsurance != null) emp.paySocialInsurance = saved.paySocialInsurance;
        if (saved.socialInsuranceBase != null) emp.socialInsuranceBase = saved.socialInsuranceBase;
      }
    });
  } catch (e) {
    console.warn('Failed to load employee cache:', e);
  }
}

function saveEmployeeCache(emp) {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    const cache = raw ? JSON.parse(raw) : {};
    const key = getCacheKey(emp);
    cache[key] = {
      monthlySalary: emp.monthlySalary,
      probationSalary: emp.probationSalary,
      probationMonths: emp.probationMonths,
      paySocialInsurance: emp.paySocialInsurance,
      socialInsuranceBase: emp.socialInsuranceBase,
    };
    localStorage.setItem(CACHE_KEY, JSON.stringify(cache));
  } catch (e) {
    console.warn('Failed to save employee cache:', e);
  }
}

function saveEmployeesCache(fileName) {
  try {
    localStorage.setItem(EMPLOYEES_CACHE_KEY, JSON.stringify(employees));
    if (fileName) localStorage.setItem(FILENAME_CACHE_KEY, fileName);
  } catch (e) {
    console.warn('Failed to save employees cache:', e);
  }
}

function loadEmployeesFromCache() {
  try {
    const raw = localStorage.getItem(EMPLOYEES_CACHE_KEY);
    if (!raw) return false;
    employees = JSON.parse(raw);
    loadEmployeeCache(employees);
    // Restore filename display
    const cachedName = localStorage.getItem(FILENAME_CACHE_KEY);
    if (cachedName) {
      document.getElementById('fileName').textContent = cachedName;
      document.getElementById('fileName').parentElement.classList.add('has-file');
    }
    return true;
  } catch (e) {
    console.warn('Failed to load employees cache:', e);
    return false;
  }
}

// ======= Utilities =======
function escHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function fmtMoney(v) {
  return '¥' + v.toFixed(2);
}

// ======= File Upload =======
document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  loadFile(file);
});

// Drag & drop
const uploadArea = document.getElementById('uploadArea');
uploadArea.addEventListener('dragover', function(e) {
  e.preventDefault();
  this.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', function() {
  this.classList.remove('dragover');
});
uploadArea.addEventListener('drop', function(e) {
  e.preventDefault();
  this.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) loadFile(file);
});

function loadFile(file) {
  document.getElementById('fileName').textContent = file.name;
  document.getElementById('fileName').parentElement.classList.add('has-file');

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      employees = parseExcel(data);
      loadEmployeeCache(employees);
      saveEmployeesCache(file.name);
      render();
    } catch (err) {
      alert('解析 Excel 失败: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ======= Export CSV =======
function exportCSV() {
  if (!employees.length) return;

  // 校验当月入职员工是否勾选了缴纳社保
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-indexed
  const currentMonthStr = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}`;

  const newHireWithInsurance = employees.filter(emp => {
    if (!emp.hireDate) return false;
    const hireMonth = emp.hireDate.substring(0, 7);
    return hireMonth === currentMonthStr && emp.paySocialInsurance !== false;
  });

  if (newHireWithInsurance.length > 0) {
    const names = newHireWithInsurance.map(e => e.name).join('、');
    const confirmed = confirm(`以下当月入职员工已勾选缴纳社保：${names}\n\n当月入职员工通常无需缴纳当月社保，请确认是否继续导出？`);
    if (!confirmed) return;
  }

  // 校验试用期员工是否勾选了缴纳社保
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const probationWithInsurance = employees.filter(emp => {
    if (!emp.hireDate || emp.paySocialInsurance === false) return false;
    const months = emp.probationMonths != null ? emp.probationMonths : probationSettings.months;
    if (months === 0) return false;
    const hire = new Date(emp.hireDate);
    const endDate = new Date(hire);
    endDate.setMonth(endDate.getMonth() + months);
    endDate.setDate(endDate.getDate() - 1);
    return today <= endDate;
  });

  if (probationWithInsurance.length > 0) {
    const names = probationWithInsurance.map(e => e.name).join('、');
    const confirmed = confirm(`以下试用期员工已勾选缴纳社保：${names}\n\n试用期员工通常无需缴纳社保，请确认是否继续导出？`);
    if (!confirmed) return;
  }

  // Calculate company contribution totals
  const si = socialInsurance;
  let csv = '﻿姓名,工号,部门,入职日期,转正日期,正式月薪,试用期月薪,试用期月数,是否缴纳社保,社保基数,应出勤天数,实际出勤天数,缺勤小时,病假小时,调休假小时,事假小时,病假扣款,事假扣款,';
  csv += `个人养老(${si.pension.personal}%),个人医疗(${si.medical.personal}%),个人失业(${si.unemployment.personal}%),个人生育(${si.maternity.personal}%),个人工伤(${si.injury.personal}%),个人社保合计,`;
  csv += `单位养老(${si.pension.unit}%),单位医疗(${si.medical.unit}%),单位失业(${si.unemployment.unit}%),单位生育(${si.maternity.unit}%),单位工伤(${si.injury.unit}%),单位社保合计,`;
  csv += '实发工资,单位社保总额,企业用人总成本\n';

  employees.forEach(emp => {
    const calc = calcEmployee(emp);
    const formalSalary = emp.monthlySalary;
    const probationMonths = emp.probationMonths != null ? emp.probationMonths : probationSettings.months;
    const probationSalary = probationMonths === 0 ? formalSalary : (emp.probationSalary || (formalSalary * probationSettings.discount / 100));
    const probationEndDateStr = calc.probationEndDate ? calc.probationEndDate.toISOString().split('T')[0] : '';
    const hireDateStr = emp.hireDate || '';

    // Company contribution (based on formal salary)
    const companyPension = formalSalary * si.pension.unit / 100;
    const companyMedical = formalSalary * si.medical.unit / 100;
    const companyUnemployment = formalSalary * si.unemployment.unit / 100;
    const companyMaternity = formalSalary * si.maternity.unit / 100;
    const companyInjury = formalSalary * si.injury.unit / 100;
    const companySiTotal = companyPension + companyMedical + companyUnemployment + companyMaternity + companyInjury;

    const totalCost = calc.netSalary + companySiTotal;

    csv += `${emp.name},${emp.employeeId},${emp.dept},${hireDateStr},${probationEndDateStr},${formalSalary},${probationSalary},${probationMonths},${emp.paySocialInsurance !== false ? '是' : '否'},${emp.socialInsuranceBase > 0 ? emp.socialInsuranceBase : formalSalary},`;
    csv += `${emp.requiredDays},${emp.actualDays},${(emp.absenceHours - emp.compHours).toFixed(1)},${emp.sickHours},${emp.compHours},${emp.personalHours},`;
    csv += `${calc.sickDeduction.toFixed(2)},${calc.personalDeduction.toFixed(2)},`;
    csv += `${calc.pensionDeduction.toFixed(2)},${calc.medicalDeduction.toFixed(2)},${calc.unemploymentDeduction.toFixed(2)},${calc.maternityDeduction.toFixed(2)},${calc.injuryDeduction.toFixed(2)},${calc.siTotal.toFixed(2)},`;
    csv += `${companyPension.toFixed(2)},${companyMedical.toFixed(2)},${companyUnemployment.toFixed(2)},${companyMaternity.toFixed(2)},${companyInjury.toFixed(2)},${companySiTotal.toFixed(2)},`;
    csv += `${calc.netSalary.toFixed(2)},${companySiTotal.toFixed(2)},${totalCost.toFixed(2)}\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = '工资计算结果.csv';
  link.click();
  URL.revokeObjectURL(link.href);
}

// ======= Init: restore from cache on page load =======
(function() {
  if (loadEmployeesFromCache()) {
    render();
  }
})();
