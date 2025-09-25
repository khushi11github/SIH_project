// timetableGenerator.js
const XLSX = require('xlsx');
const path = require('path');

class TimetableGenerator {
    constructor() {
        this.teachers = [];
        this.classes = [];
        this.subjects = [];
        this.timeSlots = [];
        this.specialPeriods = [];
        this.timetable = {};
        this.config = {};
        // Caches for speed
        this.teacherAssignments = {}; // slotKey -> teacherId
        this.teacherDailyHours = {}; // teacherId -> { [day]: count }
        this.availabilityCache = new Map(); // key: teacherId|day|start-end -> boolean
    }

    // Read data from Excel files
    async loadDataFromExcel(excelDir) {
        try {
            // Load teachers
            const teachersWorkbook = XLSX.readFile(path.join(excelDir, 'teachers.xlsx'));
            this.teachers = this.parseTeachersData(XLSX.utils.sheet_to_json(teachersWorkbook.Sheets[teachersWorkbook.SheetNames[0]]));
            
            // Load classes
            const classesWorkbook = XLSX.readFile(path.join(excelDir, 'classes.xlsx'));
            this.classes = this.parseClassesData(XLSX.utils.sheet_to_json(classesWorkbook.Sheets[classesWorkbook.SheetNames[0]]));
            
            // Load subjects
            const subjectsWorkbook = XLSX.readFile(path.join(excelDir, 'subjects.xlsx'));
            this.subjects = this.parseSubjectsData(XLSX.utils.sheet_to_json(subjectsWorkbook.Sheets[subjectsWorkbook.SheetNames[0]]));
            
            // Load config
            const configWorkbook = XLSX.readFile(path.join(excelDir, 'config.xlsx'));
            this.config = this.parseConfigData(XLSX.utils.sheet_to_json(configWorkbook.Sheets[configWorkbook.SheetNames[0]]));
            
            console.log('Data loaded successfully from Excel files');
            // Normalize data; persist only if explicitly enabled via config
            this.normalizeData();
            if (this.config.normalizeAndPersist === true) {
                await this.writeDataBackToExcel(excelDir);
            }
            return true;
        } catch (error) {
            console.error('Error loading Excel data:', error);
            return false;
        }
    }

 parseTeachersData(teachersData) {
    // Safely parse each teacher row
    return teachersData.map(teacher => ({
        id: String(teacher.id || ''),
        name: teacher.name || '',
        subjects: teacher.subjects ? teacher.subjects.split(',').map(s => String(s).trim()) : [],
        primarySubjects: teacher.primarySubjects ? teacher.primarySubjects.split(',').map(s => String(s).trim()) : [],
        availability: teacher.availability ? this.parseAvailability(teacher.availability) : [],
        maxDailyHours: Number(teacher.maxDailyHours || teacher.MaxDailyHours || 0),
        rating: Number(teacher.rating || teacher.Rating || 0)
    }));
}


    parseClassesData(classesData) {
        return classesData.map(classObj => ({
            id: String(classObj.id),
            name: classObj.name,
            room: classObj.room,
            subjects: String(classObj.subjects || '')
                .split(',')
                .map(s => String(s).trim())
                .filter(Boolean),
            totalCredits: Number(classObj.totalCredits || classObj.TotalCredits || 0)
        }));
    }

    parseSubjectsData(subjectsData) {
        const getNumber = (row, keys, fallback = 0) => {
            for (const key of keys) {
                if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
                    const num = Number(row[key]);
                    if (!Number.isNaN(num)) return num;
                }
            }
            return fallback;
        };

        return subjectsData.map(row => ({
            id: String(row.id),
            name: row.name,
            credits: getNumber(row, [
                'credits','Credits','credit','Credit','CR','cr'
            ], 1),
            weeklySessions: getNumber(row, [
                'weeklySessions','WeeklySessions','weekly_sessions','Weekly_Sessions','sessionsPerWeek','SessionsPerWeek','sessions per week','Sessions per week','Weekly Sessions','weekly sessions','PeriodsPerWeek','periodsPerWeek','periods per week','Periods per week'
            ], 1)
        }));
    }

    parseConfigData(configData) {
        const config = {};
        configData.forEach(row => {
            config[row.key] = row.value;
        });
        
        const parsed = {
            days: config.days.split(','),
            startTime: config.startTime,
            endTime: config.endTime,
            periodDuration: parseInt(config.periodDuration),
            specialPeriods: this.parseSpecialPeriods(config.specialPeriods),
            fillAllPeriods: config.fillAllPeriods === undefined ? true : String(config.fillAllPeriods).toLowerCase() === 'true'
        };
        // Optional: enable writing normalized data back to Excel
        parsed.normalizeAndPersist = config.normalizeAndPersist ? String(config.normalizeAndPersist).toLowerCase() === 'true' : false;
        parsed.branchingLimit = config.branchingLimit ? parseInt(config.branchingLimit) : 5;
        parsed.perDaySubjectCap = config.perDaySubjectCap ? parseInt(config.perDaySubjectCap) : 1;
        return parsed;
    }

   parseAvailability(availabilityStr) {
    // If the cell is missing or empty, return empty array
    if (!availabilityStr || typeof availabilityStr !== 'string') return [];

    return availabilityStr.split(',').map(avail => {
        if (!avail) return null; // skip empty entries
        const [day, timeRange] = avail.split(':');
        if (!day || !timeRange) return null; // skip invalid format
        const [startTime, endTime] = timeRange.split('-');
        if (!startTime || !endTime) return null; // skip incomplete ranges
        return {
            day: day.trim(),
            startTime: startTime.trim(),
            endTime: endTime.trim()
        };
    }).filter(entry => entry !== null); // remove null entries
}

   parseSpecialPeriods(specialPeriodsStr) {
    // If the cell is empty or missing, return empty array
    if (!specialPeriodsStr || typeof specialPeriodsStr !== 'string') return [];

    return specialPeriodsStr.split(',').map(period => {
        if (!period) return null; // skip empty entries
        const parts = period.split(':');
        if (parts.length !== 3) return null; // skip invalid format
        const [day, timeRange, type] = parts;
        if (!timeRange) return null; // skip if timeRange missing
        const timeParts = timeRange.split('-');
        if (timeParts.length !== 2) return null; // skip invalid time range
        const [startTime, endTime] = timeParts;

        return {
            day: day.trim(),
            startTime: startTime.trim(),
            endTime: endTime.trim(),
            type: type.trim()
        };
    }).filter(entry => entry !== null); // remove null entries
}


    // Initialize timetable generator
    initializeGenerator() {
        this.timeSlots = this.generateTimeSlots();
        this.initializeEmptyTimetable();
        // Reset caches
        this.teacherAssignments = {};
        this.teacherDailyHours = {};
        this.availabilityCache = new Map();
    }
    
    normalizeData() {
        // Ensure fillAllPeriods is enabled to fill whole week
        if (!this.config.fillAllPeriods) {
            this.config.fillAllPeriods = true;
        }

        // Compute daily slots per day
        const slotsPerDay = (() => {
            // Derive without regenerating slots: compute by duration and day length
            try {
                const start = this.timeToMinutes(this.config.startTime);
                const end = this.timeToMinutes(this.config.endTime);
                const dur = (this.config.periodDuration || 60) * 60;
                return Math.max(1, Math.floor((end - start) / dur));
            } catch (_) {
                return 7;
            }
        })();

        // Default teacher availability if missing: full week, full day
        this.teachers = this.teachers.map(t => {
            const hasAvail = Array.isArray(t.availability) && t.availability.length > 0;
            const availability = hasAvail ? t.availability : (this.config.days || []).map(d => ({
                day: d.trim(),
                startTime: this.config.startTime,
                endTime: this.config.endTime
            }));
            const maxDailyHours = (t.maxDailyHours && t.maxDailyHours > 0) ? t.maxDailyHours : slotsPerDay || 7;
            return { ...t, availability, maxDailyHours };
        });

        // Default subjects weekly sessions if missing: 5 (one per weekday)
        this.subjects = this.subjects.map(s => {
            const weeklySessions = (s.weeklySessions && s.weeklySessions > 0) ? s.weeklySessions : (this.config.days ? this.config.days.length : 5);
            const credits = (s.credits && s.credits > 0) ? s.credits : 1;
            return { ...s, weeklySessions, credits };
        });

        // Ensure classes have totalCredits
        this.classes = this.classes.map(c => {
            const subjectCredits = (c.subjects || []).reduce((sum, sid) => {
                const subj = this.subjects.find(s => s.id === sid);
                return sum + (subj ? (subj.credits || 0) : 0);
            }, 0);
            const totalCredits = c.totalCredits && c.totalCredits > 0 ? c.totalCredits : subjectCredits;
            return { ...c, totalCredits };
        });
    }

    async writeDataBackToExcel(excelDir) {
        try {
            // Write teachers
            const teachersSheetData = this.teachers.map(t => ({
                id: t.id,
                name: t.name,
                subjects: (t.subjects || []).join(', '),
                primarySubjects: (t.primarySubjects || []).join(', '),
                availability: (t.availability || []).map(a => `${a.day}:${a.startTime}-${a.endTime}`).join(','),
                maxDailyHours: t.maxDailyHours,
                rating: t.rating
            }));
            const teachersWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(teachersWb, XLSX.utils.json_to_sheet(teachersSheetData), 'Sheet1');
            XLSX.writeFile(teachersWb, path.join(excelDir, 'teachers.xlsx'));

            // Write subjects
            const subjectsSheetData = this.subjects.map(s => ({
                id: s.id,
                name: s.name,
                credits: s.credits,
                weeklySessions: s.weeklySessions
            }));
            const subjectsWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(subjectsWb, XLSX.utils.json_to_sheet(subjectsSheetData), 'Sheet1');
            XLSX.writeFile(subjectsWb, path.join(excelDir, 'subjects.xlsx'));

            // Write classes (preserve current structure)
            const classesSheetData = this.classes.map(c => ({
                id: c.id,
                name: c.name,
                room: c.room,
                subjects: (c.subjects || []).join(', '),
                totalCredits: c.totalCredits
            }));
            const classesWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(classesWb, XLSX.utils.json_to_sheet(classesSheetData), 'Sheet1');
            XLSX.writeFile(classesWb, path.join(excelDir, 'classes.xlsx'));

            // Write config with fillAllPeriods
            const configRows = [
                { key: 'days', value: (this.config.days || []).join(',') },
                { key: 'startTime', value: this.config.startTime },
                { key: 'endTime', value: this.config.endTime },
                { key: 'periodDuration', value: this.config.periodDuration },
                { key: 'specialPeriods', value: (this.config.specialPeriods || []).map(sp => `${sp.day}:${sp.startTime}-${sp.endTime}:${sp.type}`).join(',') },
                { key: 'fillAllPeriods', value: String(this.config.fillAllPeriods) }
            ];
            const configWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(configWb, XLSX.utils.json_to_sheet(configRows), 'Sheet1');
            XLSX.writeFile(configWb, path.join(excelDir, 'config.xlsx'));
        } catch (e) {
            console.warn('Warning: Failed to write back to Excel files. Continuing with in-memory data.', e.message);
        }
    }
validateData() {
    const errors = [];

    // Check if classes have subjects
    this.classes.forEach(classObj => {
        if (!classObj.subjects || classObj.subjects.length === 0) {
            errors.push(`Class ${classObj.name} has no subjects assigned`);
        }
    });

    // Check if subjects have teachers
    this.subjects.forEach(subject => {
        const teachersForSubject = this.teachers.filter(teacher => 
            teacher.subjects.includes(subject.id)
        );
        if (teachersForSubject.length === 0) {
            errors.push(`Subject ${subject.name} has no qualified teachers`);
        }
    });

    // Check teacher availability
    this.teachers.forEach(teacher => {
        if (teacher.maxDailyHours <= 0) {
            errors.push(`Teacher ${teacher.name} has invalid max daily hours`);
        }
    });

    if (errors.length > 0) {
        console.error('Data validation errors:');
        errors.forEach(error => console.error(`- ${error}`));
        return false;
    }

    console.log('Data validation passed');
    return true;
}
    generateTimeSlots() {
        const slots = [];
        const { days, startTime, endTime, periodDuration } = this.config;

        days.forEach(day => {
            let currentTime = this.timeToMinutes(startTime);
            const endTimeMinutes = this.timeToMinutes(endTime);
            
            while (currentTime < endTimeMinutes) {
                const slotStart = this.minutesToTime(currentTime);
                const slotEnd = this.minutesToTime(currentTime + periodDuration * 60);
                
                slots.push({
                    day: day,
                    startTime: slotStart,
                    endTime: slotEnd,
                    isSpecialPeriod: false,
                    specialType: null
                });
                
                currentTime += periodDuration * 60;
            }
        });

        return this.markSpecialPeriods(slots);
    }

    markSpecialPeriods(slots) {
        return slots.map(slot => {
            const special = this.config.specialPeriods.find(sp => 
                sp.day === slot.day && 
                sp.startTime === slot.startTime
            );
            
            if (special) {
                return {
                    ...slot,
                    isSpecialPeriod: true,
                    specialType: special.type
                };
            }
            return slot;
        });
    }

    timeToMinutes(timeStr) {
        const [hours, minutes] = timeStr.split(':').map(Number);
        return hours * 60 + minutes;
    }

    minutesToTime(minutes) {
        const hours = Math.floor(minutes / 60);
        const mins = minutes % 60;
        return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}`;
    }

    initializeEmptyTimetable() {
        this.timetable = {};
        this.classes.forEach(classObj => {
            this.timetable[classObj.id] = {
                className: classObj.name,
                room: classObj.room,
                schedule: {}
            };
            
            this.timeSlots.forEach(slot => {
                const slotKey = `${slot.day}_${slot.startTime}`;
                this.timetable[classObj.id].schedule[slotKey] = {
                    teacher: null,
                    subject: null,
                    room: classObj.room,
                    isSpecialPeriod: slot.isSpecialPeriod,
                    specialType: slot.specialType
                };
            });
        });
    }

    // Core timetable generation algorithm
    generateTimetable() {
        // Greedy pre-fill to quickly cover most slots
        this.greedyPrefill();
        const unassignedSlots = this.getUnassignedSlots();
        return this.backtrack(unassignedSlots);
    }

    greedyPrefill() {
        const days = this.config.days || [];
        // For each class and day, assign top priority subjects and available teachers first
        this.classes.forEach(classObj => {
            days.forEach(day => {
                const slots = this.timeSlots.filter(s => s.day === day);
                for (const slot of slots) {
                    const slotKey = `${slot.day}_${slot.startTime}`;
                    const current = this.timetable[classObj.id].schedule[slotKey];
                    if (current.isSpecialPeriod || current.subject) continue;
                    const candidates = this.getAvailableAssignments({ classId: classObj.id, slot: slot, slotKey, class: classObj, room: classObj.room }, classObj);
                    for (const assignment of candidates) {
                        if (this.isConsistent(assignment, { classId: classObj.id, slot: slot, slotKey })) {
                            // Assign and update caches
                            this.timetable[classObj.id].schedule[slotKey] = {
                                teacher: assignment.teacher,
                                subject: assignment.subject,
                                room: classObj.room,
                                isSpecialPeriod: false,
                                specialType: null
                            };
                            this.teacherAssignments[slotKey] = assignment.teacher.id;
                            const dailyMap = this.teacherDailyHours[assignment.teacher.id] || {};
                            dailyMap[slot.day] = (dailyMap[slot.day] || 0) + 1;
                            this.teacherDailyHours[assignment.teacher.id] = dailyMap;
                            break;
                        }
                    }
                }
            });
        });
    }

    getUnassignedSlots() {
        const unassigned = [];
        
        this.classes.forEach(classObj => {
            this.timeSlots.forEach(slot => {
                const slotKey = `${slot.day}_${slot.startTime}`;
                const currentAssignment = this.timetable[classObj.id].schedule[slotKey];
                
                if (!currentAssignment.teacher && !slot.isSpecialPeriod) {
                    unassigned.push({
                        classId: classObj.id,
                        class: classObj,
                        slot: slot,
                        slotKey: slotKey,
                        room: classObj.room
                    });
                }
            });
        });

        return this.sortByPriority(unassigned);
    }

    sortByPriority(unassignedSlots) {
        return unassignedSlots.sort((a, b) => {
            const classA = a.class;
            const classB = b.class;
            
            const creditDiff = classB.totalCredits - classA.totalCredits;
            if (creditDiff !== 0) return creditDiff;
            
            return classB.subjects.length - classA.subjects.length;
        });
    }

    backtrack(unassignedSlots) {
        if (unassignedSlots.length === 0) {
            return true;
        }

        // Interleave across days by selecting next slot from the least-filled day for this class
        const [currentSlot, remainingSlots] = this.pickNextSlotMRV(unassignedSlots);
        const classObj = currentSlot.class;

        const availableAssignments = this.getAvailableAssignments(currentSlot, classObj);

        // Try all valid teacher/subject assignments first
        for (const assignment of availableAssignments) {
            if (this.isConsistent(assignment, currentSlot)) {
                this.timetable[currentSlot.classId].schedule[currentSlot.slotKey] = {
                    teacher: assignment.teacher,
                    subject: assignment.subject,
                    room: currentSlot.room,
                    isSpecialPeriod: false,
                    specialType: null
                };
                // Update caches
                const slotKey = `${currentSlot.slot.day}_${currentSlot.slot.startTime}`;
                this.teacherAssignments[slotKey] = assignment.teacher.id;
                const dailyMap = this.teacherDailyHours[assignment.teacher.id] || {};
                dailyMap[currentSlot.slot.day] = (dailyMap[currentSlot.slot.day] || 0) + 1;
                this.teacherDailyHours[assignment.teacher.id] = dailyMap;

                if (this.backtrack(remainingSlots)) {
                    return true;
                }

                this.timetable[currentSlot.classId].schedule[currentSlot.slotKey] = {
                    teacher: null,
                    subject: null,
                    room: currentSlot.room,
                    isSpecialPeriod: false,
                    specialType: null
                };
                // Revert caches
                delete this.teacherAssignments[slotKey];
                const revertMap = this.teacherDailyHours[assignment.teacher.id] || {};
                if (revertMap[currentSlot.slot.day]) {
                    revertMap[currentSlot.slot.day] = Math.max(0, revertMap[currentSlot.slot.day] - 1);
                    this.teacherDailyHours[assignment.teacher.id] = revertMap;
                }
            }
        }

        // Leave this period free and proceed as a fallback to avoid dead-ends
        this.timetable[currentSlot.classId].schedule[currentSlot.slotKey] = {
            teacher: null,
            subject: null,
            room: currentSlot.room,
            isSpecialPeriod: false,
            specialType: null
        };

        if (this.backtrack(remainingSlots)) {
            return true;
        }

        // Revert (not strictly necessary since nulls, but keep symmetry)
        this.timetable[currentSlot.classId].schedule[currentSlot.slotKey] = {
            teacher: null,
            subject: null,
            room: currentSlot.room,
            isSpecialPeriod: false,
            specialType: null
        };

        return false;
    }

    pickNextSlotInterleaved(unassignedSlots) {
        if (unassignedSlots.length <= 1) return [unassignedSlots[0], unassignedSlots.slice(1)];
        // Group by class and day counts to pick the class/day with the least usage
        const usageKey = slot => `${slot.classId}_${slot.slot.day}`;
        const usageCount = {};
        unassignedSlots.forEach(s => {
            const key = usageKey(s);
            if (!(key in usageCount)) usageCount[key] = 0;
        });
        // Count already assigned non-free slots per class/day
        Object.keys(this.timetable).forEach(classId => {
            const schedule = this.timetable[classId].schedule;
            Object.entries(schedule).forEach(([slotKey, assignment]) => {
                const [day] = slotKey.split('_');
                const key = `${classId}_${day}`;
                if (!(key in usageCount)) usageCount[key] = 0;
                if (assignment && assignment.subject) usageCount[key] += 1;
            });
        });

        let bestIdx = 0;
        let bestScore = Infinity;
        for (let i = 0; i < unassignedSlots.length; i++) {
            const s = unassignedSlots[i];
            const key = usageKey(s);
            const score = usageCount[key] ?? 0;
            if (score < bestScore) {
                bestScore = score;
                bestIdx = i;
            }
        }
        const picked = unassignedSlots[bestIdx];
        const remaining = unassignedSlots.slice(0, bestIdx).concat(unassignedSlots.slice(bestIdx + 1));
        return [picked, remaining];
    }

    // Minimum Remaining Values (MRV): choose slot with the fewest available assignments
    pickNextSlotMRV(unassignedSlots) {
        if (unassignedSlots.length <= 1) return [unassignedSlots[0], unassignedSlots.slice(1)];
        let bestIdx = 0;
        let bestCount = Infinity;
        for (let i = 0; i < unassignedSlots.length; i++) {
            const s = unassignedSlots[i];
            const classObj = s.class;
            const count = this.getAvailableAssignmentsCount(s, classObj);
            if (count < bestCount) {
                bestCount = count;
                bestIdx = i;
            }
        }
        const picked = unassignedSlots[bestIdx];
        const remaining = unassignedSlots.slice(0, bestIdx).concat(unassignedSlots.slice(bestIdx + 1));
        return [picked, remaining];
    }

    getAvailableAssignmentsCount(slot, classObj) {
        const prioritySubjects = this.getPrioritySubjects(classObj);
        let count = 0;
        for (const subject of prioritySubjects) {
            const availableTeachers = this.getAvailableTeachers(subject.id, slot);
            if (availableTeachers.length > 0) count += availableTeachers.length;
            if (count > 3) return 4; // early exit cap
        }
        return count;
    }

    getAvailableAssignments(slot, classObj) {
        const assignments = [];
        const prioritySubjects = this.getPrioritySubjects(classObj);

        prioritySubjects.forEach(subject => {
            const availableTeachers = this.getAvailableTeachers(subject.id, slot);
            
            availableTeachers.forEach(teacher => {
                assignments.push({
                    teacher: teacher,
                    subject: subject,
                    priority: this.calculatePriorityScore(subject, teacher, classObj)
                });
            });
        });

        // If we need to fill all periods, allow repeating subjects beyond weeklySessions
        // but still keep one-per-day cap in consistency check
        const sorted = assignments.sort((a, b) => b.priority - a.priority);
        const limit = Math.max(1, Math.min(this.config.branchingLimit || 3, sorted.length));
        if (this.config.fillAllPeriods) {
            // Limit branching factor to speed up backtracking
            return sorted.slice(0, limit);
        }
        // Otherwise, prefer subjects that still have remaining sessions first
        const byRemaining = sorted.sort((a, b) => {
            const remA = Math.max(0, (a.subject.weeklySessions || 0) - this.getWeeklySubjectCount(classObj.id, a.subject.id));
            const remB = Math.max(0, (b.subject.weeklySessions || 0) - this.getWeeklySubjectCount(classObj.id, b.subject.id));
            return remB - remA;
        });
        return byRemaining.slice(0, limit);
    }

    calculatePriorityScore(subject, teacher, classObj) {
        let score = subject.credits * 10;
        score += teacher.rating * 2;
        
        if (teacher.primarySubjects.includes(subject.id)) {
            score += 5;
        }
        
        return score;
    }

    getPrioritySubjects(classObj) {
        return this.subjects
            .filter(subject => classObj.subjects.includes(subject.id))
            .sort((a, b) => b.credits - a.credits);
    }

    getAvailableTeachers(subjectId, slot) {
        return this.teachers.filter(teacher => {
            if (!teacher.subjects.includes(subjectId)) return false;
            if (!this.isTeacherAvailable(teacher.id, slot)) return false;
            return true;
        });
    }

    isTeacherAvailable(teacherId, slot) {
        const slotKey = `${slot.slot.day}_${slot.slot.startTime}`;
        if (this.teacherAssignments[slotKey] === teacherId) return false;
        // Fallback to scan in case cache missing
        for (const classId in this.timetable) {
            const assignment = this.timetable[classId].schedule[slotKey];
            if (assignment.teacher && assignment.teacher.id === teacherId) {
                return false;
            }
        }
        const teacher = this.teachers.find(t => t.id === teacherId);
        return this.isTeacherAvailableAtTime(teacher, slot.slot);
    }

  isTeacherAvailableAtTime(teacher, slot) {
    // If teacher has no specific availability, assume they're available
    if (!teacher.availability || teacher.availability.length === 0) {
        return true;
    }

    const cacheKey = `${teacher.id}|${slot.day}|${slot.startTime}-${slot.endTime}`;
    if (this.availabilityCache.has(cacheKey)) return this.availabilityCache.get(cacheKey);

    const slotDayAvailability = teacher.availability.find(avail => 
        avail.day.toLowerCase() === slot.day.toLowerCase()
    );
    if (!slotDayAvailability) {
        this.availabilityCache.set(cacheKey, false);
        return false; // Teacher not available on this day
    }
    const slotStart = this.timeToMinutes(slot.startTime);
    const slotEnd = this.timeToMinutes(slot.endTime);
    const availStart = this.timeToMinutes(slotDayAvailability.startTime);
    const availEnd = this.timeToMinutes(slotDayAvailability.endTime);
    const ok = slotStart >= availStart && slotEnd <= availEnd;
    this.availabilityCache.set(cacheKey, ok);
    return ok;
}

    isConsistent(assignment, slot) {
    const teacher = assignment.teacher;
    const subject = assignment.subject;
    const classId = slot.classId;
    const day = slot.slot.day;

    // Check teacher availability more rigorously
    if (!this.isTeacherAvailableAtTime(teacher, slot.slot)) {
        return false;
    }

    // Check if teacher is already assigned elsewhere at this time
    if (this.isTeacherAlreadyAssigned(teacher.id, slot.slot)) {
        return false;
    }

    // Check teacher's daily workload
    if (!this.checkTeacherWorkload(teacher, day)) {
        return false;
    }

    // Check subject distribution for the class (weekly and daily)
    if (!this.checkSubjectDistribution(classId, subject, day)) {
        return false;
    }

    // Check if teacher is qualified for this subject
    if (!teacher.subjects.includes(subject.id)) {
        return false;
    }

    return true;
}

    isTeacherAlreadyAssigned(teacherId, slot) {
        const slotKey = `${slot.day}_${slot.startTime}`;
        if (this.teacherAssignments[slotKey] === teacherId) return true;
        for (const classId in this.timetable) {
            const assignment = this.timetable[classId].schedule[slotKey];
            if (assignment.teacher && assignment.teacher.id === teacherId) {
                return true;
            }
        }
        return false;
    }

    checkTeacherWorkload(teacher, day) {
        // If maxDailyHours is not set or <= 0, treat as unlimited
        if (!teacher.maxDailyHours || teacher.maxDailyHours <= 0) {
            return true;
        }
        const dailyMap = this.teacherDailyHours[teacher.id] || {};
        const dailyHours = dailyMap[day] || 0;
        return dailyHours < teacher.maxDailyHours;
    }

    getTeacherDailyHours(teacherId, day) {
        const dailyMap = this.teacherDailyHours[teacherId] || {};
        return dailyMap[day] || 0;
    }

    checkSubjectDistribution(classId, subject, day) {
        // Enforce per-day cap per subject per class (configurable)
        const cap = Math.max(1, this.config.perDaySubjectCap || 1);
        const dailyCount = this.getDailySubjectCount(classId, subject.id, day);
        if (dailyCount >= cap) return false;

        // If we want to fill all periods, ignore weekly caps entirely
        if (this.config.fillAllPeriods) {
            return true;
        }

        const weeklySubjectCount = this.getWeeklySubjectCount(classId, subject.id);
        const hasWeeklyCap = subject.weeklySessions && subject.weeklySessions > 0;
        if (hasWeeklyCap && weeklySubjectCount >= subject.weeklySessions) {
            return false;
        }
        return true;
    }

    getWeeklySubjectCount(classId, subjectId) {
        let count = 0;
        Object.values(this.timetable[classId].schedule).forEach(assignment => {
            if (assignment.subject && assignment.subject.id === subjectId) {
                count++;
            }
        });
        return count;
    }

    getDailySubjectCount(classId, subjectId, day) {
        let count = 0;
        Object.entries(this.timetable[classId].schedule).forEach(([slotKey, assignment]) => {
            if (!slotKey.startsWith(day + '_')) return;
            if (assignment.subject && assignment.subject.id === subjectId) {
                count++;
            }
        });
        return count;
    }

    // Compute remaining required sessions (sum of each subject's required sessions minus assigned ones) for a class
    getRemainingRequiredSessionsForClass(classId) {
        const classObj = this.classes.find(c => c.id === classId);
        if (!classObj) return 0;

        let remaining = 0;
        this.subjects
            .filter(subject => classObj.subjects.includes(subject.id))
            .forEach(subject => {
                const assigned = this.getWeeklySubjectCount(classId, subject.id);
                const needed = Math.max(0, (subject.weeklySessions || 0) - assigned);
                remaining += needed;
            });

        return remaining;
    }

    // Export results to Excel
    exportTimetableToExcel(outputPath) {
        try {
            // Create class timetables workbook
            const classWorkbook = XLSX.utils.book_new();
            
            this.classes.forEach(classObj => {
                const classData = this.formatClassTimetableForExcel(classObj.id);
                const worksheet = XLSX.utils.json_to_sheet(classData);
                XLSX.utils.book_append_sheet(classWorkbook, worksheet, classObj.name);
            });
            
            XLSX.writeFile(classWorkbook, path.join(outputPath, 'class_timetables.xlsx'));
            
            // Create teacher timetables workbook
            const teacherWorkbook = XLSX.utils.book_new();
            
            this.teachers.forEach(teacher => {
                const teacherData = this.formatTeacherTimetableForExcel(teacher.id);
                const worksheet = XLSX.utils.json_to_sheet(teacherData);
                XLSX.utils.book_append_sheet(teacherWorkbook, worksheet, teacher.name);
            });
            
            XLSX.writeFile(teacherWorkbook, path.join(outputPath, 'teacher_timetables.xlsx'));
            
            console.log('Timetables exported to Excel successfully');
            return true;
        } catch (error) {
            console.error('Error exporting to Excel:', error);
            return false;
        }
    }

    formatClassTimetableForExcel(classId) {
        const classObj = this.classes.find(c => c.id === classId);
        const days = [...new Set(this.timeSlots.map(slot => slot.day))];
        const excelData = [];
        
        days.forEach(day => {
            const daySlots = this.timeSlots.filter(slot => slot.day === day);
            
            daySlots.forEach(slot => {
                const slotKey = `${day}_${slot.startTime}`;
                const assignment = this.timetable[classId].schedule[slotKey];
                
                excelData.push({
                    Day: day,
                    Time: `${slot.startTime} - ${slot.endTime}`,
                    Subject: assignment.subject ? assignment.subject.name : assignment.specialType || 'Free',
                    Teacher: assignment.teacher ? assignment.teacher.name : 'Free',
                    Room: assignment.room,
                    Type: assignment.isSpecialPeriod ? 'Special Period' : 'Regular Class'
                });
            });
            
            // Add empty row between days for better readability
            excelData.push({});
        });
        
        return excelData;
    }

    formatTeacherTimetableForExcel(teacherId) {
        const teacher = this.teachers.find(t => t.id === teacherId);
        const days = [...new Set(this.timeSlots.map(slot => slot.day))];
        const excelData = [];
        
        days.forEach(day => {
            const daySlots = this.timeSlots.filter(slot => slot.day === day);
            
            daySlots.forEach(slot => {
                const slotKey = `${day}_${slot.startTime}`;
                let assignedClass = null;
                let subject = null;
                let room = null;
                
                for (const classId in this.timetable) {
                    const assignment = this.timetable[classId].schedule[slotKey];
                    if (assignment.teacher && assignment.teacher.id === teacherId) {
                        assignedClass = this.timetable[classId].className;
                        subject = assignment.subject ? assignment.subject.name : 'Free';
                        room = assignment.room;
                        break;
                    }
                }
                
                excelData.push({
                    Day: day,
                    Time: `${slot.startTime} - ${slot.endTime}`,
                    Class: assignedClass || 'Free',
                    Subject: subject || 'Free',
                    Room: room || 'Staff Room',
                    Status: assignedClass ? 'Teaching' : 'Free'
                });
            });
            
            excelData.push({});
        });
        
        return excelData;
    }

    // Display timetable in console
    displayTimetable() {
        const result = {};
        
        this.classes.forEach(classObj => {
            result[classObj.name] = {
                room: classObj.room,
                schedule: {}
            };
            
            const days = [...new Set(this.timeSlots.map(slot => slot.day))];
            
            days.forEach(day => {
                result[classObj.name].schedule[day] = this.timeSlots
                    .filter(slot => slot.day === day)
                    .map(slot => {
                        const slotKey = `${day}_${slot.startTime}`;
                        const assignment = this.timetable[classObj.id].schedule[slotKey];
                        
                        return {
                            time: `${slot.startTime}-${slot.endTime}`,
                            teacher: assignment.teacher ? assignment.teacher.name : 'Free',
                            subject: assignment.subject ? assignment.subject.name : 
                                    assignment.specialType || 'Free',
                            room: assignment.room
                        };
                    });
            });
        });
        
        return result;
    }
}

// Usage example
async function main() {
    const timetableGenerator = new TimetableGenerator();
    
    // Load data from Excel files
    const dataLoaded = await timetableGenerator.loadDataFromExcel('./excel_data');
    
    if (!dataLoaded) {
        console.error('Failed to load data from Excel files');
        return;
    }
    
    // Validate data before generation
    if (!timetableGenerator.validateData()) {
        console.error('Data validation failed. Please fix the issues above.');
        return;
    }
    
    // Initialize and generate timetable
    timetableGenerator.initializeGenerator();
    
    console.log('Starting timetable generation...');
    console.log(`Total classes: ${timetableGenerator.classes.length}`);
    console.log(`Total teachers: ${timetableGenerator.teachers.length}`);
    console.log(`Total time slots: ${timetableGenerator.timeSlots.length}`);
    
    const success = timetableGenerator.generateTimetable();
    
    if (success) {
        console.log('Timetable generated successfully!');
        
        // Display in console
        const formattedTimetable = timetableGenerator.displayTimetable();
        console.log('\n=== CLASS TIMETABLES ===');
        console.log(JSON.stringify(formattedTimetable, null, 2));
        
        // Export to Excel
        timetableGenerator.exportTimetableToExcel('./output');
        
    } else {
        console.log('Unable to generate timetable with given constraints.');
        console.log('This could be due to:');
        console.log('- Insufficient teacher availability');
        console.log('- Conflicting constraints');
        console.log('- Not enough time slots for all subjects');
        
        // Suggest solutions
        console.log('\nSuggested solutions:');
        console.log('1. Increase teacher availability');
        console.log('2. Reduce subject requirements');
        console.log('3. Add more time slots');
        console.log('4. Hire more teachers for overloaded subjects');
    }
}

// Run the main function only when executed directly
if (require.main === module) {
    main().catch(console.error);
}

module.exports = TimetableGenerator;