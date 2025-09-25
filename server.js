const express = require('express');
const path = require('path');
const TimetableGenerator = require('./src/timetable.js');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Generate and get timetables
app.get('/api/generate', async (req, res) => {
	try {
		const tg = new TimetableGenerator();
		const loaded = await tg.loadDataFromExcel(path.join(__dirname, 'excel_data'));
		if (!loaded || !tg.validateData()) {
			return res.status(400).json({ error: 'Invalid data' });
		}
		tg.initializeGenerator();
		const ok = tg.generateTimetable();
		if (!ok) {
			return res.status(409).json({ error: 'Unable to generate timetable with given constraints' });
		}
		// Fill activities
		tg.fillFreeSlotsWithActivities();
		const classes = tg.displayTimetable();
		return res.json({ classes });
	} catch (e) {
		return res.status(500).json({ error: e.message });
	}
});

// Class timetable by class id
app.get('/api/classes/:classId', async (req, res) => {
	try {
		const tg = new TimetableGenerator();
		const loaded = await tg.loadDataFromExcel(path.join(__dirname, 'excel_data'));
		if (!loaded || !tg.validateData()) {
			return res.status(400).json({ error: 'Invalid data' });
		}
		tg.initializeGenerator();
		if (!tg.generateTimetable()) {
			return res.status(409).json({ error: 'Unable to generate timetable with given constraints' });
		}
		const classId = req.params.classId;
		const data = tg.formatClassTimetableForExcel(classId);
		return res.json({ classId, data });
	} catch (e) {
		return res.status(500).json({ error: e.message });
	}
});

// Teacher timetable by teacher id
app.get('/api/teachers/:teacherId', async (req, res) => {
	try {
		const tg = new TimetableGenerator();
		const loaded = await tg.loadDataFromExcel(path.join(__dirname, 'excel_data'));
		if (!loaded || !tg.validateData()) {
			return res.status(400).json({ error: 'Invalid data' });
		}
		tg.initializeGenerator();
		if (!tg.generateTimetable()) {
			return res.status(409).json({ error: 'Unable to generate timetable with given constraints' });
		}
		const teacherId = req.params.teacherId;
		const data = tg.formatTeacherTimetableForExcel(teacherId);
		return res.json({ teacherId, data });
	} catch (e) {
		return res.status(500).json({ error: e.message });
	}
});

app.listen(PORT, () => {
	console.log(`Server listening on http://localhost:${PORT}`);
});
