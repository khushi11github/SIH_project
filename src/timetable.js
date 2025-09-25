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
            this.subjects = XLSX.utils.sheet_to_json(subjectsWorkbook.Sheets[subjectsWorkbook.SheetNames[0]]);
            
            // Load config
            const configWorkbook = XLSX.readFile(path.join(excelDir, 'config.xlsx'));
            this.config = this.parseConfigData(XLSX.utils.sheet_to_json(configWorkbook.Sheets[configWorkbook.SheetNames[0]]));
            
            console.log('Data loaded successfully from Excel files');
            return true;
        } catch (error) {
            console.error('Error loading Excel data:', error);
            return false;
        }
    }

 parseTeachersData(teachersData) {
    // Safely parse each teacher row
    return teachersData.map(teacher => ({
        id: teacher.id || '',
        name: teacher.name || '',
        subjects: teacher.subjects ? teacher.subjects.split(',').map(s => s.trim()) : [],
        primarySubjects: teacher.primarySubjects ? teacher.primarySubjects.split(',').map(s => s.trim()) : [],
        availability: teacher.availability ? this.parseAvailability(teacher.availability) : [],
        maxDailyHours: teacher.maxDailyHours || 0,
        rating: teacher.rating || 0
    }));
}


    parseClassesData(classesData) {
        return classesData.map(classObj => ({
            id: classObj.id,
            name: classObj.name,
            room: classObj.room,
            subjects: classObj.subjects.split(',').map(s => s.trim()),
            totalCredits: classObj.totalCredits
        }));
    }

    parseConfigData(configData) {
        const config = {};
        configData.forEach(row => {
            config[row.key] = row.value;
        });
        
        return {
            days: config.days.split(','),
            startTime: config.startTime,
            endTime: config.endTime,
            periodDuration: parseInt(config.periodDuration),
            specialPeriods: this.parseSpecialPeriods(config.specialPeriods)
        };
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
        const unassignedSlots = this.getUnassignedSlots();
        return this.backtrack(unassignedSlots);
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

        const currentSlot = unassignedSlots[0];
        const remainingSlots = unassignedSlots.slice(1);
        const classObj = currentSlot.class;

        const availableAssignments = this.getAvailableAssignments(currentSlot, classObj);

        for (const assignment of availableAssignments) {
            if (this.isConsistent(assignment, currentSlot)) {
                this.timetable[currentSlot.classId].schedule[currentSlot.slotKey] = {
                    teacher: assignment.teacher,
                    subject: assignment.subject,
                    room: currentSlot.room,
                    isSpecialPeriod: false,
                    specialType: null
                };

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
            }
        }

        return false;
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

        return assignments.sort((a, b) => b.priority - a.priority);
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
        for (const classId in this.timetable) {
            const slotKey = `${slot.slot.day}_${slot.slot.startTime}`;
            const assignment = this.timetable[classId].schedule[slotKey];
            
            if (assignment.teacher && assignment.teacher.id === teacherId) {
                return false;
            }
        }
        
        const teacher = this.teachers.find(t => t.id === teacherId);
        return this.isTeacherAvailableAtTime(teacher, slot.slot);
    }

    isTeacherAvailableAtTime(teacher, slot) {
        const teacherAvailability = teacher.availability.find(avail => 
            avail.day === slot.day
        );
        
        if (!teacherAvailability) return false;
        
        const slotStart = this.timeToMinutes(slot.startTime);
        const slotEnd = this.timeToMinutes(slot.endTime);
        const availStart = this.timeToMinutes(teacherAvailability.startTime);
        const availEnd = this.timeToMinutes(teacherAvailability.endTime);
        
        return slotStart >= availStart && slotEnd <= availEnd;
    }

    isConsistent(assignment, slot) {
        if (!this.checkTeacherWorkload(assignment.teacher, slot.slot.day)) {
            return false;
        }
        
        if (!this.checkSubjectDistribution(slot.classId, assignment.subject)) {
            return false;
        }
        
        return true;
    }

    checkTeacherWorkload(teacher, day) {
        const dailyHours = this.getTeacherDailyHours(teacher.id, day);
        return dailyHours < teacher.maxDailyHours;
    }

    getTeacherDailyHours(teacherId, day) {
        let hours = 0;
        
        for (const classId in this.timetable) {
            Object.keys(this.timetable[classId].schedule).forEach(slotKey => {
                if (slotKey.startsWith(day)) {
                    const assignment = this.timetable[classId].schedule[slotKey];
                    if (assignment.teacher && assignment.teacher.id === teacherId) {
                        hours++;
                    }
                }
            });
        }
        
        return hours;
    }

    checkSubjectDistribution(classId, subject) {
        const weeklySubjectCount = this.getWeeklySubjectCount(classId, subject.id);
        return weeklySubjectCount < subject.weeklySessions;
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
    
    // Initialize and generate timetable
    timetableGenerator.initializeGenerator();
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
    }
}

// Run the main function
main().catch(console.error);

module.exports = TimetableGenerator;