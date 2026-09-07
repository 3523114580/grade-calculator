const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const { parseGradeExcel, calculateComprehensive, rankStudents } = require('./gradeCalculator');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '请上传文件' });
    }
    const result = parseGradeExcel(req.file.buffer);
    res.json(result);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '解析失败：' + err.message });
  }
});

app.post('/api/calculate', (req, res) => {
  try {
    const { students, extras } = req.body;
    if (!students || !extras) {
      return res.status(400).json({ error: '数据不完整' });
    }

    const results = students.map(student => {
      const extra = extras[student.studentId] || {};
      const scores = calculateComprehensive(student, extra);
      return { ...student, ...scores };
    });

    const comprehensiveRanks = rankStudents(results, 'comprehensive');
    const academicRanks = rankStudents(results, 'academicScore');

    results.forEach(r => {
      r.comprehensiveRank = comprehensiveRanks[r.studentId];
      r.academicRank = academicRanks[r.studentId];
    });

    results.sort((a, b) => a.comprehensiveRank - b.comprehensiveRank);

    res.json({ results });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '计算失败：' + err.message });
  }
});

app.post('/api/export', (req, res) => {
  try {
    const { results, className } = req.body;
    if (!results) {
      return res.status(400).json({ error: '数据不完整' });
    }

    const wb = XLSX.utils.book_new();

    const summaryData = [
      ['附件4  2025-2026学年班级学习成绩与综合考评成绩排名汇总表'],
      ['学院：          专业：数据科学与大数据技术(专升本)          班级：' + (className || '') + '          日期：2026年  月  日'],
      ['序号', '学号', '姓名', '班级', '学业成绩', '学业成绩排名', '综合考评成绩', '综合考评排名', '备注']
    ];
    results.forEach((r, i) => {
      summaryData.push([
        i + 1, r.studentId, r.name, className || '',
        r.academicScore, r.academicRank,
        r.comprehensive, r.comprehensiveRank, ''
      ]);
    });
    const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
    ws1['!cols'] = [{ wch: 6 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws1, '附件4-成绩排名汇总');

    const detailData = [
      ['附件1  2025-2026学年学生综合素质评分表'],
      ['学院：    专业：数据科学与大数据技术(专升本)    班级：' + (className || '')],
      ['序号', '学号', '姓名',
        '德育素养(100分;20%)', '', '',
        '智育素养(100分;50%)', '',
        '体育素养(100分;10%)', '',
        '美育素养(100分;10%)', '',
        '劳动教育素养(100分;10%)', '', '',
        '总评', '名次']
    ];
    detailData.push(['', '', '',
      '基础分', '奖励分', '扣分',
      '学业成绩', '学业表现',
      '基础分', '奖励分',
      '基础分', '奖励分',
      '基础分', '奖励分', '扣分',
      '', '']);
    results.forEach((r, i) => {
      const moralReward = (r.extras && r.extras.moralReward) || 0;
      const moralDeduct = (r.extras && r.extras.moralDeduct) || 0;
      const academicPerformance = (r.extras && r.extras.academicPerformance) || 0;
      const sportsBase = (r.extras && r.extras.sportsBase !== undefined) ? r.extras.sportsBase : 60;
      const sportsReward = (r.extras && r.extras.sportsReward) || 0;
      const aestheticBase = (r.extras && r.extras.aestheticBase !== undefined) ? r.extras.aestheticBase : 60;
      const aestheticReward = (r.extras && r.extras.aestheticReward) || 0;
      const laborReward = (r.extras && r.extras.laborReward) || 0;
      const laborDeduct = (r.extras && r.extras.laborDeduct) || 0;

      detailData.push([
        i + 1, r.studentId, r.name,
        60, moralReward, moralDeduct,
        r.intellectualAcademicPart, academicPerformance,
        sportsBase, sportsReward,
        aestheticBase, aestheticReward,
        60, laborReward, laborDeduct,
        r.comprehensive, r.comprehensiveRank
      ]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(detailData);
    XLSX.utils.book_append_sheet(wb, ws2, '附件1-综合素质评分表');

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + encodeURIComponent('综合素质评价结果.xlsx'));
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '导出失败：' + err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`成绩计算系统已启动: http://localhost:${PORT}`);
});
