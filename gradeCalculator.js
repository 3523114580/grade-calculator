const XLSX = require('xlsx');

function parseGradeExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const courseRow = data[5];
  const creditRow = data[6];

  const courses = [];
  for (let col = 4; col < courseRow.length; col++) {
    const courseName = String(courseRow[col] || '').trim();
    const credit = parseFloat(creditRow[col]);
    if (courseName && !isNaN(credit) && credit > 0) {
      courses.push({ col, name: courseName, credit });
    }
  }

  const students = [];
  for (let row = 7; row < data.length; row++) {
    const rowData = data[row];
    const studentId = String(rowData[1] || '').trim();
    const studentName = String(rowData[2] || '').trim();
    if (!studentId || !studentName) continue;

    const courseScores = {};
    let totalScoreCredit = 0;
    let totalCredit = 0;
    let courseCount = 0;

    courses.forEach(course => {
      const raw = rowData[course.col];
      let score = null;
      if (raw !== '' && raw !== null && raw !== undefined) {
        score = parseFloat(raw);
        if (isNaN(score)) score = null;
      }
      courseScores[course.name] = score;
      if (score !== null && score >= 0) {
        totalScoreCredit += score * course.credit;
        totalCredit += course.credit;
        courseCount++;
      }
    });

    const academicScore = totalCredit > 0 ? (totalScoreCredit / totalCredit) : 0;
    const intellectualAcademicPart = academicScore * 0.8;

    students.push({
      index: parseInt(rowData[0]) || students.length + 1,
      studentId,
      name: studentName,
      courseScores,
      courseCount,
      totalCredit,
      academicScore: round2(academicScore),
      intellectualAcademicPart: round2(intellectualAcademicPart)
    });
  }

  return { courses, students };
}

function round2(num) {
  return Math.round(num * 100) / 100;
}

function calculateComprehensive(student, extra) {
  const moralBase = 60;
  const moralReward = parseFloat(extra.moralReward) || 0;
  const moralDeduct = parseFloat(extra.moralDeduct) || 0;
  const F1 = clamp(moralBase + moralReward - moralDeduct, 0, 100);

  const academicPart = student.intellectualAcademicPart;
  const academicPerformance = parseFloat(extra.academicPerformance) || 0;
  const F2 = clamp(academicPart + academicPerformance, 0, 100);

  const sportsBase = parseFloat(extra.sportsBase);
  const sportsBaseFinal = isNaN(sportsBase) ? 60 : sportsBase;
  const sportsReward = parseFloat(extra.sportsReward) || 0;
  const F3 = clamp(sportsBaseFinal + sportsReward, 0, 100);

  const aestheticBase = parseFloat(extra.aestheticBase);
  const aestheticBaseFinal = isNaN(aestheticBase) ? 60 : aestheticBase;
  const aestheticReward = parseFloat(extra.aestheticReward) || 0;
  const F4 = clamp(aestheticBaseFinal + aestheticReward, 0, 100);

  const laborBase = 60;
  const laborReward = parseFloat(extra.laborReward) || 0;
  const laborDeduct = parseFloat(extra.laborDeduct) || 0;
  const F5 = clamp(laborBase + laborReward - laborDeduct, 0, 100);

  const F = F1 * 0.2 + F2 * 0.5 + F3 * 0.1 + F4 * 0.1 + F5 * 0.1;

  return {
    F1: round2(F1),
    F2: round2(F2),
    F3: round2(F3),
    F4: round2(F4),
    F5: round2(F5),
    comprehensive: round2(F)
  };
}

function clamp(val, min, max) {
  return Math.max(min, Math.min(max, val));
}

function rankStudents(studentsWithScores, field) {
  const sorted = [...studentsWithScores].sort((a, b) => b[field] - a[field]);
  const ranks = {};
  sorted.forEach((s, i) => {
    ranks[s.studentId] = i + 1;
  });
  return ranks;
}

module.exports = {
  parseGradeExcel,
  calculateComprehensive,
  rankStudents,
  round2
};
