import React, { useState, useEffect } from 'react';
import { AlertTriangle, CheckCircle, Download, Info, RotateCcw, Trash2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import './App.css';

const PhysicistScheduler = () => {
    const [physicists, setPhysicists] = useState([]);
    const [availability, setAvailability] = useState({});
    const [dutyCaps, setDutyCaps] = useState({});
    const [schedule, setSchedule] = useState({});
    const [currentWeek, setCurrentWeek] = useState(0);
    const [weeks, setWeeks] = useState([]);
    const [startDate, setStartDate] = useState('2025-06-09');
    const [numberOfWeeks, setNumberOfWeeks] = useState(4);
    const [viewMode, setViewMode] = useState('week');
    const [currentPhysicist, setCurrentPhysicist] = useState(0);
    const [dataStatus, setDataStatus] = useState({
        capsLoaded: false,
        availabilityLoaded: false
    });
    const [undoStack, setUndoStack] = useState([]);
    const [showConflictModal, setShowConflictModal] = useState(false);
    const [conflictData, setConflictData] = useState(null);
    const [showDutyConflictModal, setShowDutyConflictModal] = useState(false);
    const [dutyConflictData, setDutyConflictData] = useState(null);
    const [showExportModal, setShowExportModal] = useState(false);
    const [exportContent, setExportContent] = useState('');

    const duties = [
        'Unity DPOD1', 'Unity DPOD2', 'Unity POD1', 'Unity POD2', 'Unity Backup',
        'AROC DPOD', 'AROC TPOD Early', 'AROC TPOD Late', 'Ethos TPOD1', 'Ethos TPOD2',
        'Ethos Planning1', 'Ethos Planning2', 'ART Float', 'EROC DPOD', 'EROC TPOD Early',
        'EROC TPOD Late', 'BgRT POD', 'AROC HDR Physics', 'AROC HDR Backup', 'CUH HDR',
        'LDR', 'GK', 'CK', 'CK Backup', 'GPOD', 'IVBT', 'VMAT TBI'
    ];

    const weeklyDuties = [
        'Unity DPOD1', 'Unity DPOD2', 'Unity Backup', 'Ethos Planning1', 'Ethos Planning2', 'ART Float', 'CK', 'CK Backup'
    ];

    const combinedDutyMap = {
        'Unity DPOD1': 'Unity DPOD',
        'Unity DPOD2': 'Unity DPOD',
        'Unity POD1': 'Unity POD',
        'Unity POD2': 'Unity POD',
        'AROC TPOD Early': 'AROC TPOD',
        'AROC TPOD Late': 'AROC TPOD',
        'Ethos TPOD1': 'Ethos TPOD',
        'Ethos TPOD2': 'Ethos TPOD',
        'Ethos Planning1': 'Ethos Planning',
        'Ethos Planning2': 'Ethos Planning',
        'EROC TPOD Early': 'EROC TPOD',
        'EROC TPOD Late': 'EROC TPOD'
    };

    const pairedDuties = {
        'Unity DPOD1': ['Unity DPOD1', 'Unity DPOD2'],
        'Unity DPOD2': ['Unity DPOD1', 'Unity DPOD2'],
        'Unity POD1': ['Unity POD1', 'Unity POD2'],
        'Unity POD2': ['Unity POD1', 'Unity POD2'],
        'AROC TPOD Early': ['AROC TPOD Early', 'AROC TPOD Late'],
        'AROC TPOD Late': ['AROC TPOD Early', 'AROC TPOD Late'],
        'Ethos TPOD1': ['Ethos TPOD1', 'Ethos TPOD2'],
        'Ethos TPOD2': ['Ethos TPOD1', 'Ethos TPOD2'],
        'Ethos Planning1': ['Ethos Planning1', 'Ethos Planning2'],
        'Ethos Planning2': ['Ethos Planning1', 'Ethos Planning2'],
        'EROC TPOD Early': ['EROC TPOD Early', 'EROC TPOD Late'],
        'EROC TPOD Late': ['EROC TPOD Early', 'EROC TPOD Late']
    };

    const allowedCombinations = [
        ['Unity Backup', 'AROC HDR Backup'],
        ['CK Backup', 'EROC DPOD'],
        ['AROC HDR Backup', 'EROC DPOD']
    ];

    const blockedCombinations = [
        ['CK', 'CK Backup']
    ];

    const restrictedDuties = {
        'GPOD': [1, 3, 5],
        'AROC HDR Physics': [2, 4]
    };

    const nationalHolidays = [
        '2025-01-01', '2025-01-20', '2025-05-26', '2025-06-19', '2025-07-04',
        '2025-09-01', '2025-11-27', '2025-12-25'
    ];

    useEffect(() => {
        generateWeeks();
    }, [startDate, numberOfWeeks]);

    const generateWeeks = () => {
        const start = new Date(startDate + 'T12:00:00');
        const dayOfWeek = start.getDay();
        if (dayOfWeek !== 1) {
            console.warn(`Start date ${startDate} is not a Monday (day of week: ${dayOfWeek}).`);
        }

        const weekList = [];
        for (let i = 0; i < numberOfWeeks; i++) {
            const weekDates = [];
            for (let j = 0; j < 5; j++) {
                const date = new Date(start);
                date.setDate(start.getDate() + (i * 7) + j);
                weekDates.push(date);
            }
            weekList.push(weekDates);
        }
        setWeeks(weekList);

        if (currentWeek >= numberOfWeeks) {
            setCurrentWeek(0);
        }
    };

    const excelSerialToDate = (serial) => {
        const baseDate = new Date(1899, 11, 30);
        const resultDate = new Date(baseDate.getTime() + serial * 24 * 60 * 60 * 1000);
        return resultDate.toISOString().split('T')[0];
    };

    const handleDataFile = (file) => {
        if (!file) {
            alert('No file selected');
            return;
        }

        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            alert('Please upload an Excel file (.xlsx or .xls)');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                if (!workbook.SheetNames.length) {
                    alert('No sheets found in Excel file');
                    return;
                }

                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                if (!jsonData.length) {
                    alert('First sheet is empty');
                    return;
                }

                const caps = {};
                const physicistData = [];

                jsonData.forEach(row => {
                    let name = null;
                    const nameColumns = ['Physicist', 'Name', 'Clinical duties', Object.keys(row)[0]];

                    for (const col of nameColumns) {
                        if (row[col] && typeof row[col] === 'string' && row[col].trim()) {
                            name = row[col].trim();
                            break;
                        }
                    }

                    if (name && !name.toLowerCase().includes('total')) {
                        caps[name] = {};

                        duties.forEach(duty => {
                            caps[name][duty] = 0;
                        });

                        let totalSum = 0;
                        const sumColumns = ['Sum', 'sum', 'Total', 'total', 'TOTAL', 'SUM'];
                        for (const sumCol of sumColumns) {
                            if (row[sumCol] && !isNaN(parseInt(row[sumCol]))) {
                                totalSum = parseInt(row[sumCol]);
                                break;
                            }
                        }

                        if (totalSum === 0) {
                            Object.keys(row).forEach(colName => {
                                if (colName !== Object.keys(row)[0]) {
                                    const value = row[colName];
                                    if (value && !isNaN(parseInt(value))) {
                                        totalSum += parseInt(value);
                                    }
                                }
                            });
                        }

                        Object.keys(row).forEach(colName => {
                            if (colName !== Object.keys(row)[0]) {
                                const value = row[colName];
                                if (value && !isNaN(parseInt(value))) {
                                    const capValue = parseInt(value);
                                    const colLower = colName.toLowerCase().replace(/\s+/g, '');

                                    if (colLower.includes('aroc') && colLower.includes('tpod')) {
                                        caps[name]['AROC TPOD Early'] = capValue;
                                        caps[name]['AROC TPOD Late'] = capValue;
                                    } else if (colLower.includes('eroc') && colLower.includes('tpod')) {
                                        caps[name]['EROC TPOD Early'] = capValue;
                                        caps[name]['EROC TPOD Late'] = capValue;
                                    } else if (colLower.includes('unity') && colLower.includes('dpod')) {
                                        caps[name]['Unity DPOD1'] = capValue;
                                        caps[name]['Unity DPOD2'] = capValue;
                                    } else if (colLower.includes('unity') && colLower.includes('pod') && !colLower.includes('dpod')) {
                                        caps[name]['Unity POD1'] = capValue;
                                        caps[name]['Unity POD2'] = capValue;
                                    } else if (colLower.includes('ethos') && colLower.includes('tpod')) {
                                        caps[name]['Ethos TPOD1'] = capValue;
                                        caps[name]['Ethos TPOD2'] = capValue;
                                    } else if (colLower.includes('ethos') && colLower.includes('planning')) {
                                        caps[name]['Ethos Planning1'] = capValue;
                                        caps[name]['Ethos Planning2'] = capValue;
                                    } else if (duties.includes(colName)) {
                                        caps[name][colName] = capValue;
                                    }
                                }
                            }
                        });

                        physicistData.push({ name, totalSum });
                    }
                });

                physicistData.sort((a, b) => b.totalSum - a.totalSum);
                const physicistNames = physicistData.map(p => p.name);

                const avail = {};
                if (workbook.SheetNames.length > 1) {
                    const availSheet = workbook.Sheets[workbook.SheetNames[1]];
                    const availRawData = XLSX.utils.sheet_to_json(availSheet, { header: 1 });

                    if (availRawData.length > 0) {
                        for (let rowIndex = 1; rowIndex < availRawData.length; rowIndex++) {
                            const row = availRawData[rowIndex];
                            if (!row || row.length === 0) continue;

                            let physicistName = row[0];
                            if (!physicistName || physicistName.toString().trim() === '') continue;

                            physicistName = physicistName.toString().trim();
                            const hard = [];

                            for (let colIndex = 1; colIndex < row.length; colIndex++) {
                                const serialValue = row[colIndex];

                                if (serialValue !== undefined && serialValue !== null && serialValue !== '') {
                                    try {
                                        const serialNumber = parseInt(serialValue);
                                        if (!isNaN(serialNumber) && serialNumber > 1) {
                                            const isoDate = excelSerialToDate(serialNumber);
                                            hard.push(isoDate);
                                        }
                                    } catch (e) {
                                        console.log(`Error processing serial number ${serialValue}:`, e);
                                    }
                                }
                            }

                            if (hard.length > 0) {
                                avail[physicistName] = { hard, soft: [] };
                            }
                        }
                        setDataStatus(prev => ({ ...prev, availabilityLoaded: true }));
                    }
                }

                if (workbook.SheetNames.length > 3) {
                    const avoidanceSheet = workbook.Sheets[workbook.SheetNames[3]];
                    const avoidanceRawData = XLSX.utils.sheet_to_json(avoidanceSheet, { header: 1 });

                    if (avoidanceRawData.length > 0) {
                        for (let rowIndex = 1; rowIndex < avoidanceRawData.length; rowIndex++) {
                            const row = avoidanceRawData[rowIndex];
                            if (!row || row.length === 0) continue;

                            let physicistName = row[0];
                            if (!physicistName || physicistName.toString().trim() === '') continue;

                            physicistName = physicistName.toString().trim();

                            if (!avail[physicistName]) {
                                avail[physicistName] = { hard: [], soft: [] };
                            }
                            if (!avail[physicistName].soft) {
                                avail[physicistName].soft = [];
                            }

                            for (let colIndex = 1; colIndex < row.length; colIndex++) {
                                const serialValue = row[colIndex];

                                if (serialValue !== undefined && serialValue !== null && serialValue !== '') {
                                    try {
                                        const serialNumber = parseInt(serialValue);
                                        if (!isNaN(serialNumber) && serialNumber > 1) {
                                            const isoDate = excelSerialToDate(serialNumber);

                                            if (!avail[physicistName].soft.includes(isoDate)) {
                                                avail[physicistName].soft.push(isoDate);
                                            }
                                        }
                                    } catch (e) {
                                        console.log(`Error processing avoidance serial number ${serialValue}:`, e);
                                    }
                                }
                            }
                        }
                    }
                }

                setPhysicists(physicistNames);
                setDutyCaps(caps);
                setAvailability(avail);
                setDataStatus(prev => ({ ...prev, capsLoaded: true }));

                alert(`Successfully loaded ${physicistNames.length} physicists!`);

            } catch (error) {
                console.error('Error processing Excel file:', error);
                alert('Error processing Excel file: ' + error.message);
            }
        };

        reader.readAsArrayBuffer(file);
    };

    const getCurrentAssignments = (physicist) => {
        const assignments = {};
        duties.forEach(duty => {
            assignments[duty] = 0;
        });

        Object.values(schedule).forEach(weekSchedule => {
            if (weekSchedule) {
                Object.entries(weekSchedule).forEach(([dateStr, daySchedule]) => {
                    if (daySchedule) {
                        Object.entries(daySchedule).forEach(([duty, assignedPhysicist]) => {
                            if (assignedPhysicist === physicist) {
                                if (pairedDuties[duty]) {
                                    pairedDuties[duty].forEach(pairedDuty => {
                                        assignments[pairedDuty] = (assignments[pairedDuty] || 0) + 1;
                                    });
                                } else {
                                    assignments[duty] = (assignments[duty] || 0) + 1;
                                }
                            }
                        });
                    }
                });
            }
        });

        return assignments;
    };

    const handleAssignment = (duty, date, weekIndex, physicist) => {
        const dateStr = date.toISOString().split('T')[0];
        const daySchedule = schedule[weekIndex]?.[dateStr] || {};
        const existingAssignment = Object.entries(daySchedule).find(([existingDuty, assignedPhysicist]) =>
            assignedPhysicist === physicist && existingDuty !== duty
        );
        if (physicist && existingAssignment) {
            const [existingDuty] = existingAssignment;
            const canCombine = allowedCombinations.some(combo =>
                combo.includes(duty) && combo.includes(existingDuty)
            );
            if (!canCombine) {
                setConflictData({
                    physicist,
                    existingDuty,
                    newDuty: duty,
                    date,
                    weekIndex,
                    dateStr
                });
                setShowConflictModal(true);
                return;
            }
        }
        performAssignment(duty, date, weekIndex, physicist);
    };

    const handlePhysicistAssignment = (duty, date, weekIndex, physicist) => {
        const dateStr = date.toISOString().split('T')[0];
        const daySchedule = schedule[weekIndex]?.[dateStr] || {};
        const currentlyAssigned = daySchedule[duty];
        if (currentlyAssigned && currentlyAssigned !== physicist) {
            setDutyConflictData({
                duty,
                date,
                weekIndex,
                dateStr,
                currentlyAssigned,
                newPhysicist: physicist
            });
            setShowDutyConflictModal(true);
            return;
        }
        handleAssignment(duty, date, weekIndex, physicist);
    };

    const handleDutyConflictResolution = (shouldReplace) => {
        setShowDutyConflictModal(false);
        if (shouldReplace) {
            handleAssignment(
                dutyConflictData.duty,
                dutyConflictData.date,
                dutyConflictData.weekIndex,
                dutyConflictData.newPhysicist
            );
        }
        setDutyConflictData(null);
    };

    const handleConflictResolution = (shouldSwitch) => {
        setShowConflictModal(false);
        if (shouldSwitch) {
            performAssignment(
                conflictData.newDuty,
                conflictData.date,
                conflictData.weekIndex,
                conflictData.physicist,
                conflictData.existingDuty
            );
        } else {
            performAssignment(
                conflictData.newDuty,
                conflictData.date,
                conflictData.weekIndex,
                conflictData.physicist
            );
        }
        setConflictData(null);
    };

    const performAssignment = (duty, date, weekIndex, physicist, dutyToRemoveFrom = null) => {
        const dateStr = date.toISOString().split('T')[0];

        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        setSchedule(prev => {
            const newSchedule = { ...prev };

            if (dutyToRemoveFrom) {
                if (newSchedule[weekIndex]?.[dateStr]?.[dutyToRemoveFrom]) {
                    newSchedule[weekIndex] = {
                        ...newSchedule[weekIndex],
                        [dateStr]: {
                            ...newSchedule[weekIndex][dateStr]
                        }
                    };
                    delete newSchedule[weekIndex][dateStr][dutyToRemoveFrom];
                }
            }

            if (!newSchedule[weekIndex]) {
                newSchedule[weekIndex] = {};
            }
            if (!newSchedule[weekIndex][dateStr]) {
                newSchedule[weekIndex][dateStr] = {};
            }

            newSchedule[weekIndex][dateStr][duty] = physicist || undefined;

            return newSchedule;
        });

        if (physicist && weeklyDuties.includes(duty)) {
            const weekDates = weeks[weekIndex];

            setSchedule(prev => {
                const newSchedule = { ...prev };

                weekDates.forEach(weekDate => {
                    const weekDateStr = weekDate.toISOString().split('T')[0];
                    const dayOfWeek = weekDate.getDay();

                    if (dayOfWeek === 0 || dayOfWeek === 6) return;

                    if (nationalHolidays.includes(weekDateStr)) return;

                    const isRestricted = restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek);
                    if (isRestricted) return;

                    const isHardUnavailable = availability[physicist]?.hard?.includes(weekDateStr) &&
                        !availability[physicist]?.soft?.includes(weekDateStr);
                    if (isHardUnavailable) return;

                    const daySchedule = newSchedule[weekIndex]?.[weekDateStr] || {};
                    const hasConflict = Object.entries(daySchedule).some(([existingDuty, existingPhysicist]) => {
                        if (existingPhysicist === physicist && existingDuty !== duty) {
                            const canCombine = allowedCombinations.some(combo =>
                                combo.includes(duty) && combo.includes(existingDuty)
                            );
                            return !canCombine;
                        }
                        return false;
                    });

                    if (hasConflict) return;

                    if (!newSchedule[weekIndex]) {
                        newSchedule[weekIndex] = {};
                    }
                    if (!newSchedule[weekIndex][weekDateStr]) {
                        newSchedule[weekIndex][weekDateStr] = {};
                    }

                    newSchedule[weekIndex][weekDateStr][duty] = physicist;
                });

                return newSchedule;
            });
        }
    };

    const getWeeklyDutyAssignments = (physicist, weekIndex, duty) => {
        if (!weeklyDuties.includes(duty)) return 0;

        const weekDates = weeks[weekIndex];
        if (!weekDates) return 0;

        let assignmentCount = 0;

        weekDates.forEach(date => {
            const dateStr = date.toISOString().split('T')[0];
            const dayOfWeek = date.getDay();

            if (dayOfWeek === 0 || dayOfWeek === 6) return;

            if (nationalHolidays.includes(dateStr)) return;

            const daySchedule = schedule[weekIndex]?.[dateStr] || {};
            if (daySchedule[duty] === physicist) {
                assignmentCount++;
            }
        });

        return assignmentCount;
    };

    const wasOverAssignedPreviousWeek = (physicist, duty, currentWeekIndex) => {
        if (!weeklyDuties.includes(duty) || currentWeekIndex === 0) return false;

        const previousWeekIndex = currentWeekIndex - 1;

        if (pairedDuties[duty]) {
            const totalAssignments = pairedDuties[duty].reduce((total, pairedDuty) => {
                if (weeklyDuties.includes(pairedDuty)) {
                    return total + getWeeklyDutyAssignments(physicist, previousWeekIndex, pairedDuty);
                }
                return total;
            }, 0);

            return totalAssignments >= 3;
        } else {
            const assignmentCount = getWeeklyDutyAssignments(physicist, previousWeekIndex, duty);
            return assignmentCount >= 3;
        }
    };

    const canAssign = (physicist, duty, date, weekIndex) => {
        const dateStr = date.toISOString().split('T')[0];
        const dayOfWeek = date.getDay();

        let cap = dutyCaps[physicist]?.[duty] || 0;

        if (cap === 0 && combinedDutyMap[duty]) {
            const combinedDutyName = combinedDutyMap[duty];
            cap = dutyCaps[physicist]?.[combinedDutyName] || 0;
        }

        if (cap === 0) {
            return { canAssign: false, reason: 'No capacity (0 cap)' };
        }

        const hasSoftAvoidance = availability[physicist]?.soft?.includes(dateStr);
        const hasHardUnavailability = availability[physicist]?.hard?.includes(dateStr);

        if (nationalHolidays.includes(dateStr)) {
            return { canAssign: false, reason: 'National holiday' };
        }

        if (hasHardUnavailability && !hasSoftAvoidance) {
            return { canAssign: false, reason: 'Hard unavailable' };
        }

        if (restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek)) {
            return { canAssign: false, reason: 'Restricted day' };
        }

        const warnings = [];
        const currentAssignments = getCurrentAssignments(physicist);
        const currentCount = currentAssignments[duty] || 0;

        if (currentCount >= cap) {
            warnings.push(`Over cap (${currentCount}/${cap})`);
        }

        const daySchedule = schedule[weekIndex]?.[dateStr] || {};
        const dayAssignments = Object.entries(daySchedule)
            .filter(([d, p]) => p === physicist)
            .map(([d]) => d);

        if (dayAssignments.length > 0) {
            const canCombine = allowedCombinations.some(combo =>
                combo.includes(duty) && dayAssignments.some(assigned => combo.includes(assigned))
            );

            if (!canCombine) {
                warnings.push('Double booked');
            }
        }

        return {
            canAssign: true,
            warnings: warnings,
            isAvoidance: hasSoftAvoidance,
            hasIssues: warnings.length > 0
        };
    };

    const getEligiblePhysicists = (duty, date, weekIndex) => {
        const eligible = [];
        const ineligible = [];

        physicists.forEach(physicist => {
            const check = canAssign(physicist, duty, date, weekIndex);
            const currentAssignments = getCurrentAssignments(physicist);
            const currentCount = currentAssignments[duty] || 0;

            let cap = dutyCaps[physicist]?.[duty] || 0;
            if (cap === 0 && combinedDutyMap[duty]) {
                cap = dutyCaps[physicist]?.[combinedDutyMap[duty]] || 0;
            }

            const weekDates = weeks[weekIndex];
            let daysAvailable = 0;
            let avoidanceDays = 0;

            weekDates.forEach(weekDate => {
                const weekDateStr = weekDate.toISOString().split('T')[0];
                const dayOfWeek = weekDate.getDay();

                if (dayOfWeek === 0 || dayOfWeek === 6) return;

                if (restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek)) return;

                if (nationalHolidays.includes(weekDateStr)) return;

                const hasHardUnavailability = availability[physicist]?.hard?.includes(weekDateStr);
                const hasSoftAvoidance = availability[physicist]?.soft?.includes(weekDateStr);

                if (hasHardUnavailability && !hasSoftAvoidance) return;

                daysAvailable++;

                if (hasSoftAvoidance) {
                    avoidanceDays++;
                }
            });

            const physicistInfo = {
                name: physicist,
                currentCount,
                cap,
                utilization: cap > 0 ? currentCount / cap : 0,
                warnings: check.warnings || [],
                hasIssues: check.hasIssues || false,
                daysAvailable,
                avoidanceDays
            };

            if (check.canAssign) {
                eligible.push({
                    ...physicistInfo,
                    isAvoidance: check.isAvoidance
                });
            } else {
                ineligible.push({
                    ...physicistInfo,
                    reason: check.reason || ''
                });
            }
        });

        eligible.sort((a, b) => {
            const aRemaining = a.cap - a.currentCount;
            const bRemaining = b.cap - b.currentCount;

            if (aRemaining !== bRemaining) {
                return bRemaining - aRemaining;
            }

            if (a.hasIssues !== b.hasIssues) {
                return a.hasIssues ? 1 : -1;
            }

            if (a.isAvoidance !== b.isAvoidance) {
                return a.isAvoidance ? 1 : -1;
            }

            return a.name.localeCompare(b.name);
        });

        ineligible.sort((a, b) => a.name.localeCompare(b.name));

        return { eligible, ineligible };
    };

    const clearPhysicistWeek = (physicist, weekIndex) => {
        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        setSchedule(prev => {
            const newSchedule = { ...prev };
            const weekDates = weeks[weekIndex];

            weekDates.forEach(date => {
                const dateStr = date.toISOString().split('T')[0];
                if (newSchedule[weekIndex]?.[dateStr]) {
                    const daySchedule = newSchedule[weekIndex][dateStr];
                    const dutiesToRemove = [];

                    Object.entries(daySchedule).forEach(([duty, assignedPhysicist]) => {
                        if (assignedPhysicist === physicist) {
                            dutiesToRemove.push(duty);
                        }
                    });

                    if (dutiesToRemove.length > 0) {
                        newSchedule[weekIndex] = {
                            ...newSchedule[weekIndex],
                            [dateStr]: {
                                ...newSchedule[weekIndex][dateStr]
                            }
                        };

                        dutiesToRemove.forEach(duty => {
                            delete newSchedule[weekIndex][dateStr][duty];
                        });

                        if (Object.keys(newSchedule[weekIndex][dateStr]).length === 0) {
                            delete newSchedule[weekIndex][dateStr];
                        }
                    }
                }
            });

            return newSchedule;
        });
    };

    const clearPhysicistAllAvoidanceDays = (physicistName) => {
        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        let totalClearedCount = 0;

        setSchedule(prev => {
            const newSchedule = { ...prev };

            weeks.forEach((weekDates, weekIndex) => {
                weekDates.forEach(date => {
                    const dateStr = date.toISOString().split('T')[0];
                    const hasSoftAvoidance = availability[physicistName]?.soft?.includes(dateStr);

                    if (hasSoftAvoidance && newSchedule[weekIndex]?.[dateStr]) {
                        const daySchedule = newSchedule[weekIndex][dateStr];
                        const dutiesToRemove = [];

                        Object.entries(daySchedule).forEach(([duty, assignedPhysicist]) => {
                            if (assignedPhysicist === physicistName) {
                                dutiesToRemove.push(duty);
                            }
                        });

                        if (dutiesToRemove.length > 0) {
                            newSchedule[weekIndex] = {
                                ...newSchedule[weekIndex],
                                [dateStr]: {
                                    ...newSchedule[weekIndex][dateStr]
                                }
                            };

                            dutiesToRemove.forEach(duty => {
                                delete newSchedule[weekIndex][dateStr][duty];
                                totalClearedCount++;
                            });

                            if (Object.keys(newSchedule[weekIndex][dateStr]).length === 0) {
                                delete newSchedule[weekIndex][dateStr];
                            }
                        }
                    }
                });
            });

            return newSchedule;
        });

        if (totalClearedCount > 0) {
            alert(`Cleared ${totalClearedCount} assignments from all avoidance days for ${physicistName}`);
        } else {
            alert(`No assignments found on avoidance days for ${physicistName}`);
        }
    };

    const clearPhysicistAvoidanceDays = (physicistName, weekIndex) => {
        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        const weekDates = weeks[weekIndex];
        let clearedCount = 0;

        setSchedule(prev => {
            const newSchedule = { ...prev };

            weekDates.forEach(date => {
                const dateStr = date.toISOString().split('T')[0];
                const hasSoftAvoidance = availability[physicistName]?.soft?.includes(dateStr);

                if (hasSoftAvoidance && newSchedule[weekIndex]?.[dateStr]) {
                    const daySchedule = newSchedule[weekIndex][dateStr];
                    const dutiesToRemove = [];

                    Object.entries(daySchedule).forEach(([duty, assignedPhysicist]) => {
                        if (assignedPhysicist === physicistName) {
                            dutiesToRemove.push(duty);
                        }
                    });

                    if (dutiesToRemove.length > 0) {
                        newSchedule[weekIndex] = {
                            ...newSchedule[weekIndex],
                            [dateStr]: {
                                ...newSchedule[weekIndex][dateStr]
                            }
                        };

                        dutiesToRemove.forEach(duty => {
                            delete newSchedule[weekIndex][dateStr][duty];
                            clearedCount++;
                        });

                        if (Object.keys(newSchedule[weekIndex][dateStr]).length === 0) {
                            delete newSchedule[weekIndex][dateStr];
                        }
                    }
                }
            });

            return newSchedule;
        });

        if (clearedCount > 0) {
            alert(`Cleared ${clearedCount} assignments from avoidance days for ${physicistName} in Week ${weekIndex + 1}`);
        } else {
            alert(`No assignments found on avoidance days for ${physicistName} in Week ${weekIndex + 1}`);
        }
    };

    const clearRow = (duty) => {
        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        setSchedule(prev => {
            const newSchedule = { ...prev };
            const weekDates = weeks[currentWeek];

            weekDates.forEach(date => {
                const dateStr = date.toISOString().split('T')[0];
                if (newSchedule[currentWeek]?.[dateStr]?.[duty]) {
                    newSchedule[currentWeek] = {
                        ...newSchedule[currentWeek],
                        [dateStr]: {
                            ...newSchedule[currentWeek][dateStr]
                        }
                    };
                    delete newSchedule[currentWeek][dateStr][duty];

                    if (Object.keys(newSchedule[currentWeek][dateStr]).length === 0) {
                        delete newSchedule[currentWeek][dateStr];
                    }
                }
            });

            return newSchedule;
        });
    };

    const undoLastChange = () => {
        if (undoStack.length > 0) {
            const previousState = undoStack[undoStack.length - 1];
            setSchedule(previousState);
            setUndoStack(prev => prev.slice(0, -1));
        }
    };

    const exportSchedule = () => {
        try {
            console.log('Excel export started...');

            const allDates = [];
            const columnHeaders = [''];

            weeks.forEach((week, weekIndex) => {
                const monday = new Date(week[0]);

                for (let dayOffset = 0; dayOffset < 7; dayOffset++) {
                    const currentDate = new Date(monday);
                    currentDate.setDate(monday.getDate() + dayOffset);

                    const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
                    const dayName = dayNames[dayOffset];

                    const columnName = weekIndex === 0 ? dayName : `${dayName}_${weekIndex}`;

                    columnHeaders.push(columnName);
                    allDates.push({
                        date: currentDate,
                        dateStr: currentDate.toISOString().split('T')[0],
                        columnName: columnName,
                        weekIndex: weekIndex,
                        dayOfWeek: dayOffset
                    });
                }
            });

            const excelData = [];

            const dateRow = [''];
            allDates.forEach(dateInfo => {
                const formattedDate = `${dateInfo.date.getMonth() + 1}/${dateInfo.date.getDate()}/${dateInfo.date.getFullYear()}`;
                dateRow.push(formattedDate);
            });
            excelData.push(dateRow);

            const dayRow = [''];
            allDates.forEach(dateInfo => {
                const dayAbbrevs = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
                dayRow.push(dayAbbrevs[dateInfo.dayOfWeek]);
            });
            excelData.push(dayRow);

            duties.forEach(duty => {
                const dutyRow = [duty];

                allDates.forEach(dateInfo => {
                    const daySchedule = schedule[dateInfo.weekIndex]?.[dateInfo.dateStr] || {};
                    const assignedPhysicist = daySchedule[duty] || '';
                    dutyRow.push(assignedPhysicist);
                });

                excelData.push(dutyRow);
            });

            console.log('Excel data prepared:', excelData.length, 'rows');

            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.aoa_to_sheet(excelData);

            const colWidths = [{ width: 25 }];
            for (let i = 1; i < columnHeaders.length; i++) {
                colWidths.push({ width: 12 });
            }
            worksheet['!cols'] = colWidths;

            const sheetName = 'Physicist Schedule';
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

            const excelBuffer = XLSX.write(workbook, {
                type: 'array',
                bookType: 'xlsx'
            });

            console.log('Excel file generated, size:', excelBuffer.byteLength || excelBuffer.length, 'bytes');

            const blob = new Blob([excelBuffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            if (weeks && weeks.length > 0) {
                const endDate = weeks[weeks.length - 1][4].toISOString().split('T')[0];
                const filename = `physicist_schedule_${startDate.replace(/-/g, '')}_to_${endDate.replace(/-/g, '')}.xlsx`;

                const reader = new FileReader();
                reader.onload = () => {
                    const dataUrl = reader.result;
                    console.log('Excel file ready for download via modal');
                    setExportContent(dataUrl);
                    setShowExportModal(true);
                };
                reader.onerror = (error) => {
                    console.error('FileReader error:', error);
                    alert('Failed to prepare Excel file for download');
                };
                reader.readAsDataURL(blob);
            }

        } catch (error) {
            console.error('Excel export error:', error);
            alert('Excel export failed: ' + error.message);
        }
    };

    const autofillSchedule = () => {
        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        const newSchedule = {};
        const tempAssignments = {};
        const lastWeeklyAssignments = {};

        physicists.forEach(physicist => {
            tempAssignments[physicist] = {};
            duties.forEach(duty => {
                tempAssignments[physicist][duty] = 0;
            });
        });

        weeks.forEach((weekDates, weekIndex) => {
            newSchedule[weekIndex] = {};
            lastWeeklyAssignments[weekIndex] = {};

            weeklyDuties.forEach(duty => {
                const availableDates = weekDates.filter(date => {
                    const dayOfWeek = date.getDay();
                    const dateStr = date.toISOString().split('T')[0];

                    if (dayOfWeek === 0 || dayOfWeek === 6) return false;
                    if (restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek)) return false;
                    if (nationalHolidays.includes(dateStr)) return false;

                    return true;
                });

                if (availableDates.length === 0) return;

                const physicistScores = physicists.map(physicist => {
                    const currentAssignments = { ...getCurrentAssignments(physicist) };
                    Object.entries(tempAssignments[physicist]).forEach(([tempDuty, count]) => {
                        if (pairedDuties[tempDuty]) {
                            pairedDuties[tempDuty].forEach(pairedDuty => {
                                currentAssignments[pairedDuty] = (currentAssignments[pairedDuty] || 0) + count;
                            });
                        } else {
                            currentAssignments[tempDuty] = (currentAssignments[tempDuty] || 0) + count;
                        }
                    });

                    const currentCount = currentAssignments[duty] || 0;
                    let cap = dutyCaps[physicist]?.[duty] || 0;
                    if (cap === 0 && combinedDutyMap[duty]) {
                        cap = dutyCaps[physicist]?.[combinedDutyMap[duty]] || 0;
                    }

                    if (cap === 0 || currentCount >= cap) return { physicist, score: -1000, canAssign: false };

                    if (weeklyDuties.includes(duty) && weekIndex > 0) {
                        const lastWeekDuty = lastWeeklyAssignments[weekIndex - 1]?.[duty];
                        if (lastWeekDuty === physicist) {
                            return { physicist, score: -999, canAssign: false };
                        }
                    }

                    let score = 0;
                    const remainingCapacity = cap - currentCount;
                    const utilization = currentCount / cap;

                    score += remainingCapacity * 200;

                    if (utilization < 0.3) {
                        score += 150;
                    } else if (utilization > 0.8) {
                        score -= 100;
                    }

                    const totalAssignments = Object.values(currentAssignments).reduce((sum, count) => sum + count, 0);
                    const avgAssignments = physicists.reduce((sum, p) => {
                        const pAssignments = getCurrentAssignments(p);
                        return sum + Object.values(pAssignments).reduce((s, c) => s + c, 0);
                    }, 0) / physicists.length;

                    if (totalAssignments < avgAssignments) {
                        score += 100;
                    }

                    let canAssignAllDates = true;
                    let totalAvoidanceScore = 0;

                    availableDates.forEach(date => {
                        const dateStr = date.toISOString().split('T')[0];
                        const isHardUnavailable = availability[physicist]?.hard?.includes(dateStr) &&
                            !availability[physicist]?.soft?.includes(dateStr);
                        if (isHardUnavailable) {
                            canAssignAllDates = false;
                        } else if (availability[physicist]?.soft?.includes(dateStr)) {
                            totalAvoidanceScore -= 500;
                        }
                    });

                    if (!canAssignAllDates) {
                        return { physicist, score: -1000, canAssign: false };
                    }

                    score += totalAvoidanceScore;

                    return { physicist, score, canAssign: true };
                }).filter(item => item.canAssign).sort((a, b) => b.score - a.score);

                if (physicistScores.length > 0) {
                    const bestPhysicist = physicistScores[0].physicist;
                    lastWeeklyAssignments[weekIndex][duty] = bestPhysicist;

                    if (physicistScores[0].score > -1000) {
                        availableDates.forEach(date => {
                            const dateStr = date.toISOString().split('T')[0];

                            if (!newSchedule[weekIndex][dateStr]) {
                                newSchedule[weekIndex][dateStr] = {};
                            }
                            newSchedule[weekIndex][dateStr][duty] = bestPhysicist;

                            if (pairedDuties[duty]) {
                                pairedDuties[duty].forEach(pairedDuty => {
                                    tempAssignments[bestPhysicist][pairedDuty]++;
                                });
                            } else {
                                tempAssignments[bestPhysicist][duty]++;
                            }
                        });
                    }
                }
            });

            const dailyDuties = duties.filter(duty => !weeklyDuties.includes(duty));

            weekDates.forEach((date, dayIndex) => {
                const dayOfWeek = date.getDay();
                const dateStr = date.toISOString().split('T')[0];

                if (dayOfWeek === 0 || dayOfWeek === 6) return;
                if (nationalHolidays.includes(dateStr)) return;

                if (!newSchedule[weekIndex][dateStr]) {
                    newSchedule[weekIndex][dateStr] = {};
                }

                dailyDuties.forEach(duty => {
                    if (restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek)) return;

                    const physicistScores = physicists.map(physicist => {
                        const currentAssignments = { ...getCurrentAssignments(physicist) };
                        Object.entries(tempAssignments[physicist]).forEach(([tempDuty, count]) => {
                            if (pairedDuties[tempDuty]) {
                                pairedDuties[tempDuty].forEach(pairedDuty => {
                                    currentAssignments[pairedDuty] = (currentAssignments[pairedDuty] || 0) + count;
                                });
                            } else {
                                currentAssignments[tempDuty] = (currentAssignments[tempDuty] || 0) + count;
                            }
                        });

                        const currentCount = currentAssignments[duty] || 0;
                        let cap = dutyCaps[physicist]?.[duty] || 0;
                        if (cap === 0 && combinedDutyMap[duty]) {
                            cap = dutyCaps[physicist]?.[combinedDutyMap[duty]] || 0;
                        }

                        if (cap === 0 || currentCount >= cap) return { physicist, score: -1000, canAssign: false };

                        const isHardUnavailable = availability[physicist]?.hard?.includes(dateStr) &&
                            !availability[physicist]?.soft?.includes(dateStr);
                        if (isHardUnavailable) {
                            return { physicist, score: -1000, canAssign: false };
                        }

                        const daySchedule = newSchedule[weekIndex][dateStr] || {};
                        const existingDuties = Object.entries(daySchedule)
                            .filter(([d, p]) => p === physicist)
                            .map(([d]) => d);

                        if (existingDuties.length > 0) {
                            for (const existingDuty of existingDuties) {
                                const isBlocked = blockedCombinations.some(combo =>
                                    combo.includes(duty) && combo.includes(existingDuty)
                                );
                                if (isBlocked) {
                                    return { physicist, score: -1000, canAssign: false };
                                }

                                const isAllowed = allowedCombinations.some(combo =>
                                    combo.includes(duty) && combo.includes(existingDuty)
                                );

                                if (!isAllowed && duty !== existingDuty) {
                                    return { physicist, score: -1000, canAssign: false };
                                }
                            }
                        }

                        let score = 0;
                        const remainingCapacity = cap - currentCount;
                        const utilization = currentCount / cap;

                        score += remainingCapacity * 200;

                        if (utilization < 0.3) {
                            score += 150;
                        } else if (utilization > 0.8) {
                            score -= 100;
                        }

                        const totalAssignments = Object.values(currentAssignments).reduce((sum, count) => sum + count, 0);
                        const avgAssignments = physicists.reduce((sum, p) => {
                            const pAssignments = getCurrentAssignments(p);
                            return sum + Object.values(pAssignments).reduce((s, c) => s + c, 0);
                        }, 0) / physicists.length;

                        if (totalAssignments < avgAssignments) {
                            score += 100;
                        }

                        if (availability[physicist]?.soft?.includes(dateStr)) {
                            score -= 600;
                        }

                        return { physicist, score, canAssign: true };
                    }).filter(item => item.canAssign).sort((a, b) => b.score - a.score);

                    if (physicistScores.length > 0 && physicistScores[0].score > -500) {
                        const bestPhysicist = physicistScores[0].physicist;
                        newSchedule[weekIndex][dateStr][duty] = bestPhysicist;

                        if (pairedDuties[duty]) {
                            pairedDuties[duty].forEach(pairedDuty => {
                                tempAssignments[bestPhysicist][pairedDuty]++;
                            });
                        } else {
                            tempAssignments[bestPhysicist][duty]++;
                        }
                    }
                });
            });
        });

        setSchedule(newSchedule);
    };

    const autofillPhysicist = (physicistName) => {
        setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(schedule))]);

        const currentAssignments = getCurrentAssignments(physicistName);

        const eligibleDuties = duties.map(duty => {
            let cap = dutyCaps[physicistName]?.[duty] || 0;
            if (cap === 0 && combinedDutyMap[duty]) {
                cap = dutyCaps[physicistName]?.[combinedDutyMap[duty]] || 0;
            }

            if (cap === 0) return null;

            const currentCount = currentAssignments[duty] || 0;
            const remainingCapacity = cap - currentCount;
            const utilization = currentCount / cap;

            if (remainingCapacity <= 0) return null;

            return {
                duty,
                cap,
                currentCount,
                remainingCapacity,
                utilization
            };
        }).filter(Boolean);

        if (eligibleDuties.length === 0) {
            alert(`${physicistName} has no remaining capacity in any duties.`);
            return;
        }

        setSchedule(prev => {
            const newSchedule = JSON.parse(JSON.stringify(prev));
            let assignmentsMade = 0;
            const maxPasses = 10;

            for (let pass = 0; pass < maxPasses; pass++) {
                let assignmentsThisPass = 0;

                const sortedDuties = [...eligibleDuties]
                    .filter(d => d.remainingCapacity > 0)
                    .sort((a, b) => a.utilization - b.utilization);

                if (sortedDuties.length === 0) break;

                for (const dutyInfo of sortedDuties) {
                    const { duty } = dutyInfo;

                    if (dutyInfo.remainingCapacity <= 0) continue;

                    let bestAssignment = null;
                    let bestScore = -1000;

                    weeks.forEach((weekDates, weekIndex) => {
                        weekDates.forEach(date => {
                            const dateStr = date.toISOString().split('T')[0];
                            const dayOfWeek = date.getDay();

                            if (dayOfWeek === 0 || dayOfWeek === 6) return;
                            if (nationalHolidays.includes(dateStr)) return;

                            if (restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek)) return;

                            const isHardUnavailable = availability[physicistName]?.hard?.includes(dateStr) &&
                                !availability[physicistName]?.soft?.includes(dateStr);
                            if (isHardUnavailable) return;

                            const daySchedule = newSchedule[weekIndex]?.[dateStr] || {};

                            if (daySchedule[duty] && daySchedule[duty] !== physicistName) return;

                            if (daySchedule[duty] === physicistName) return;

                            const existingAssignments = Object.entries(daySchedule)
                                .filter(([d, p]) => p === physicistName && d !== duty)
                                .map(([d]) => d);

                            if (existingAssignments.length > 0) {
                                const hasConflict = !allowedCombinations.some(combo =>
                                    combo.includes(duty) && existingAssignments.some(assigned => combo.includes(assigned))
                                );
                                if (hasConflict) return;
                            }

                            if (weeklyDuties.includes(duty)) {
                                const validWeekDates = weekDates.filter(weekDate => {
                                    const weekDateStr = weekDate.toISOString().split('T')[0];
                                    const weekDayOfWeek = weekDate.getDay();

                                    if (weekDayOfWeek === 0 || weekDayOfWeek === 6) return false;
                                    if (nationalHolidays.includes(weekDateStr)) return false;
                                    if (restrictedDuties[duty] && !restrictedDuties[duty].includes(weekDayOfWeek)) return false;

                                    return true;
                                });

                                let canAssignWholeWeek = true;
                                for (const weekDate of validWeekDates) {
                                    const weekDateStr = weekDate.toISOString().split('T')[0];

                                    const isWeekHardUnavailable = availability[physicistName]?.hard?.includes(weekDateStr) &&
                                        !availability[physicistName]?.soft?.includes(weekDateStr);
                                    if (isWeekHardUnavailable) {
                                        canAssignWholeWeek = false;
                                        break;
                                    }

                                    const weekDaySchedule = newSchedule[weekIndex]?.[weekDateStr] || {};
                                    if (weekDaySchedule[duty] && weekDaySchedule[duty] !== physicistName) {
                                        canAssignWholeWeek = false;
                                        break;
                                    }

                                    const weekExistingAssignments = Object.entries(weekDaySchedule)
                                        .filter(([d, p]) => p === physicistName && d !== duty)
                                        .map(([d]) => d);

                                    if (weekExistingAssignments.length > 0) {
                                        let weekHasConflict = false;
                                        const isBlocked = blockedCombinations.some(combo =>
                                            combo.includes(duty) && weekExistingAssignments.some(assigned => combo.includes(assigned))
                                        );

                                        if (isBlocked) {
                                            weekHasConflict = true;
                                        } else {
                                            weekHasConflict = !allowedCombinations.some(combo =>
                                                combo.includes(duty) && weekExistingAssignments.some(assigned => combo.includes(assigned))
                                            );
                                        }

                                        if (weekHasConflict) {
                                            canAssignWholeWeek = false;
                                            break;
                                        }
                                    }
                                }

                                if (!canAssignWholeWeek) return;

                                if (dutyInfo.remainingCapacity < validWeekDates.length) return;

                                let score = 100;

                                validWeekDates.forEach(weekDate => {
                                    const weekDateStr = weekDate.toISOString().split('T')[0];
                                    if (availability[physicistName]?.soft?.includes(weekDateStr)) {
                                        score -= 400;
                                    }
                                });

                                score -= weekIndex * 5;

                                if (score > bestScore) {
                                    bestScore = score;
                                    bestAssignment = {
                                        type: 'weekly',
                                        duty,
                                        weekIndex,
                                        dates: validWeekDates
                                    };
                                }
                            } else {
                                let score = 100;

                                if (availability[physicistName]?.soft?.includes(dateStr)) {
                                    score -= 400;
                                }

                                score -= weekIndex * 5 + date.getDay();

                                if (score > bestScore) {
                                    bestScore = score;
                                    bestAssignment = {
                                        type: 'daily',
                                        duty,
                                        weekIndex,
                                        date,
                                        dateStr
                                    };
                                }
                            }
                        });
                    });

                    if (bestAssignment && bestScore > -300) {
                        if (bestAssignment.type === 'weekly') {
                            bestAssignment.dates.forEach(weekDate => {
                                const weekDateStr = weekDate.toISOString().split('T')[0];

                                if (!newSchedule[bestAssignment.weekIndex]) {
                                    newSchedule[bestAssignment.weekIndex] = {};
                                }
                                if (!newSchedule[bestAssignment.weekIndex][weekDateStr]) {
                                    newSchedule[bestAssignment.weekIndex][weekDateStr] = {};
                                }

                                newSchedule[bestAssignment.weekIndex][weekDateStr][duty] = physicistName;
                            });

                            dutyInfo.currentCount += bestAssignment.dates.length;
                            dutyInfo.remainingCapacity -= bestAssignment.dates.length;
                            dutyInfo.utilization = dutyInfo.currentCount / dutyInfo.cap;
                            assignmentsMade += bestAssignment.dates.length;
                            assignmentsThisPass++;
                        } else {
                            if (!newSchedule[bestAssignment.weekIndex]) {
                                newSchedule[bestAssignment.weekIndex] = {};
                            }
                            if (!newSchedule[bestAssignment.weekIndex][bestAssignment.dateStr]) {
                                newSchedule[bestAssignment.weekIndex][bestAssignment.dateStr] = {};
                            }

                            newSchedule[bestAssignment.weekIndex][bestAssignment.dateStr][duty] = physicistName;

                            dutyInfo.currentCount += 1;
                            dutyInfo.remainingCapacity -= 1;
                            dutyInfo.utilization = dutyInfo.currentCount / dutyInfo.cap;
                            assignmentsMade++;
                            assignmentsThisPass++;
                        }
                    }
                }

                if (assignmentsThisPass === 0) break;
            }

            if (assignmentsMade > 0) {
                alert(`Successfully made ${assignmentsMade} assignments for ${physicistName}, maintaining balance across duties`);
            } else {
                alert(`No additional assignments could be made for ${physicistName}. All constraints respected (caps, conflicts, avoidances).`);
            }

            return newSchedule;
        });
    };

    if (weeks.length === 0) return <div className="p-4">Loading...</div>;

    const isDataReady = dataStatus.capsLoaded;

    return (
        <div className="p-6 max-w-7xl mx-auto">
            <h1 className="text-3xl font-bold mb-6">Physicist Duty Scheduler</h1>

            <div className="mb-6 p-4 bg-gray-50 rounded-lg">
                <h3 className="font-semibold mb-3 flex items-center">
                    <Info className="w-4 h-4 mr-2" />
                    How to Use the Scheduler
                </h3>
                <div className="text-sm space-y-2">
                    <div><strong>Two Views Available:</strong></div>
                    <ul className="list-disc list-inside ml-4 space-y-1">
                        <li><strong>Week View:</strong> Organize by duties (rows) and days (columns) - best for filling specific duties</li>
                        <li><strong>Physicist View:</strong> Organize by individual physicists - best for balancing individual workloads</li>
                    </ul>

                    <div className="mt-3"><strong>Week View - Physicist Information Format:</strong></div>
                    <div className="ml-4">{`{Name} ({Current Assignments}/{Duty Cap}) [Days Available This Week/Avoidance Days This Week]`}</div>

                    <div className="mt-3"><strong>Physicist View Features:</strong></div>
                    <ul className="list-disc list-inside ml-4 space-y-1">
                        <li>Each physicist has their own tab showing their weekly schedule</li>
                        <li>Individual "Fill [Name]" button optimizes assignments for that specific physicist</li>
                        <li>Add duties using the dropdown in each day cell</li>
                        <li>Remove duties by clicking the × button next to assignments</li>
                        <li>Clear entire weeks using the "Clear Week" button in each week row</li>
                        <li>View assignment summary with capacity utilization at the bottom</li>
                        <li>Yellow highlighting indicates soft avoidance dates (⚠️)</li>
                        <li>Red cells indicate hard unavailability</li>
                    </ul>

                    <div className="mt-3"><strong>Visual Indicators (Both Views):</strong></div>
                    <ul className="list-disc list-inside ml-4 space-y-1">
                        <li><strong>✅</strong> - Physicist is already assigned to another duty on this date</li>
                        <li><strong>🚨</strong> - Physicist prefers to avoid this specific date</li>
                        <li><strong>⚠️</strong> - Current assignment creates a conflict (over cap or double booked)</li>
                        <li><strong>❗</strong> - Physicist was assigned to this weekly duty for 3+ days in the previous week</li>
                        <li><strong>Deep Blue Cell</strong> - Assignment violates constraints (over cap or double booking)</li>
                        <li><strong>Light Red Cell</strong> - Assignment on a soft avoidance date</li>
                    </ul>

                    <div className="mt-3"><strong>Both views stay synchronized:</strong> Changes made in either view automatically update the other!</div>
                </div>
            </div>

            <div className="mb-6 p-4 bg-green-50 rounded-lg">
                <h3 className="font-semibold mb-3 flex items-center">
                    <CheckCircle className="w-4 h-4 mr-2 text-green-600" />
                    Auto-Fill Systems
                </h3>
                <div className="text-sm space-y-2">
                    <div><strong>Global Auto-Fill (creates an optimized initial schedule by):</strong></div>
                    <ul className="list-disc list-inside ml-4 space-y-1">
                        <li><strong>Respecting all hard constraints:</strong> Duty caps, hard unavailability, restricted days, holidays</li>
                        <li><strong>Strictly preventing double booking:</strong> No physicist assigned to multiple duties on same date (except for specific allowed combinations)</li>
                        <li><strong>Aggressively avoiding soft avoidances:</strong> Very strong penalties for scheduling on avoidance dates (especially weekly duties)</li>
                        <li><strong>Strictly preventing consecutive weekly duties:</strong> Nearly impossible for physicist to get same weekly duty two weeks in a row</li>
                        <li><strong>Limiting cap overages:</strong> Maximum 4 assignments over duty cap, with escalating penalties</li>
                        <li><strong>Balancing workload:</strong> Distributes assignments evenly across all physicists throughout the quarter</li>
                        <li><strong>Smart prioritization:</strong> Weekly duties filled first (more constrained), then daily duties</li>
                        <li><strong>Capacity optimization:</strong> Prioritizes physicists with remaining capacity in their duty caps</li>
                        <li><strong>Holiday handling:</strong> Skips holidays when assigning weekly duties (4-day weeks when holidays occur)</li>
                    </ul>

                    <div className="mt-4"><strong>Individual Physicist Auto-Fill (Physicist View "Fill [Name]" button):</strong></div>
                    <ul className="list-disc list-inside ml-4 space-y-1">
                        <li><strong>Maximizes individual utilization:</strong> Prioritizes duties where the physicist has the lowest utilization rate</li>
                        <li><strong>Respects existing assignments:</strong> Works around duties already assigned to others</li>
                        <li><strong>Follows all constraints:</strong> Hard unavailability, conflicts, day restrictions, holidays</li>
                        <li><strong>Smart weekly duty handling:</strong> Assigns weekly duties to entire weeks when possible</li>
                        <li><strong>Capacity-focused:</strong> Aims to fill up remaining capacity in all duties efficiently</li>
                        <li><strong>Conflict avoidance:</strong> Skips dates that would create double-booking issues</li>
                    </ul>

                    <div className="mt-2 p-2 bg-yellow-50 border border-yellow-200 rounded">
                        <strong>Allowed Double Bookings:</strong> Unity Backup + AROC HDR Backup, CK Backup + EROC DPOD, AROC HDR Backup + EROC DPOD
                    </div>
                    <div className="mt-2 p-2 bg-red-50 border border-red-200 rounded">
                        <strong>Strict Constraints:</strong> Consecutive weekly duties nearly impossible (-999 penalty), soft avoidances heavily penalized (up to -1600 for multiple), cap overages limited to 4 max
                    </div>
                    <div className="mt-2 text-green-700"><strong>Tips:</strong> Use global auto-fill for initial schedule, then individual auto-fill to optimize specific physicists' workloads!</div>
                </div>
            </div>

            <div className="mb-6 p-4 bg-gray-50 rounded-lg">
                <h3 className="font-semibold mb-3 flex items-center">
                    <Info className="w-4 h-4 mr-2" />
                    Data Status
                </h3>
                <div className="grid grid-cols-1 gap-4">
                    <div className={`p-3 rounded ${dataStatus.capsLoaded ? 'bg-green-100' : 'bg-red-100'}`}>
                        <div className="flex items-center">
                            {dataStatus.capsLoaded ?
                                <CheckCircle className="w-4 h-4 text-green-600 mr-2" /> :
                                <AlertTriangle className="w-4 h-4 text-red-600 mr-2" />
                            }
                            <span className="font-medium">Data Loaded</span>
                        </div>
                        <div className="text-sm mt-1">
                            {dataStatus.capsLoaded ?
                                `${Object.keys(dutyCaps).length} physicists with duty caps loaded` :
                                'Upload Excel file to begin scheduling'
                            }
                        </div>
                    </div>
                </div>
            </div>

            <div className="mb-6 grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-4">
                    <h3 className="font-semibold mb-2">Upload Excel File</h3>
                    <input
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => e.target.files[0] && handleDataFile(e.target.files[0])}
                        className="w-full"
                    />
                    <div className="text-sm text-gray-600 mt-1">
                        Upload Excel file with:<br />
                        • Sheet 1: Physicist names + duty caps (required)<br />
                        • Sheet 2: Hard unavailability dates (optional)<br />
                        • Sheet 4: Soft avoidance dates (optional)
                    </div>
                </div>

                <div className="border-2 border-gray-300 rounded-lg p-4">
                    <h3 className="font-semibold mb-3">Schedule Configuration</h3>
                    <div className="space-y-3">
                        <div>
                            <label className="block text-sm font-medium mb-1">Start Date (Monday)</label>
                            <input
                                type="date"
                                value={startDate}
                                onChange={(e) => {
                                    const selectedDate = new Date(e.target.value + 'T12:00:00');
                                    const dayOfWeek = selectedDate.getDay();

                                    if (dayOfWeek !== 1) {
                                        const mondayDate = new Date(selectedDate);
                                        let daysToAdd;
                                        if (dayOfWeek === 0) {
                                            daysToAdd = 1;
                                        } else {
                                            daysToAdd = 8 - dayOfWeek;
                                        }
                                        mondayDate.setDate(selectedDate.getDate() + daysToAdd);
                                        setStartDate(mondayDate.toISOString().split('T')[0]);
                                    } else {
                                        setStartDate(e.target.value);
                                    }
                                }}
                                className="w-full p-2 border border-gray-300 rounded"
                            />
                            <div className="text-xs text-gray-500 mt-1">
                                Will automatically adjust to nearest Monday if another day is selected
                            </div>
                        </div>
                        <div>
                            <label className="block text-sm font-medium mb-1">Number of Weeks (1-13)</label>
                            <input
                                type="number"
                                min="1"
                                max="13"
                                value={numberOfWeeks}
                                onChange={(e) => setNumberOfWeeks(Math.min(13, Math.max(1, parseInt(e.target.value) || 1)))}
                                className="w-full p-2 border border-gray-300 rounded"
                            />
                        </div>
                        <div className="text-sm text-gray-600">
                            Schedule will run from {startDate} for {numberOfWeeks} week{numberOfWeeks !== 1 ? 's' : ''}
                            {weeks.length > 0 && (
                                <div className="mt-1">
                                    ({weeks[0][0].toLocaleDateString()} - {weeks[weeks.length - 1][4].toLocaleDateString()})
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            </div>

            {isDataReady && Object.keys(schedule).length === 0 && (
                <div className="mb-4">
                    <button
                        onClick={autofillSchedule}
                        className="bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600 flex items-center"
                    >
                        <CheckCircle className="w-4 h-4 mr-2" />
                        Generate Initial Schedule
                    </button>
                </div>
            )}

            {Object.keys(schedule).length > 0 && isDataReady && (
                <div className="mb-4 flex gap-2">
                    <button
                        onClick={autofillSchedule}
                        className="bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600 flex items-center"
                    >
                        <CheckCircle className="w-4 h-4 mr-2" />
                        Auto-Fill Schedule
                    </button>
                    <button
                        onClick={exportSchedule}
                        className="bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600 flex items-center"
                    >
                        <Download className="w-4 h-4 mr-2" />
                        Export Schedule (Excel)
                    </button>
                    <button
                        onClick={undoLastChange}
                        disabled={undoStack.length === 0}
                        className="bg-gray-500 text-white px-4 py-2 rounded hover:bg-gray-600 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center"
                    >
                        <RotateCcw className="w-4 h-4 mr-2" />
                        Undo ({undoStack.length})
                    </button>
                </div>
            )}

            {isDataReady && (
                <>
                    <div className="mb-6">
                        <div className="flex space-x-1 bg-gray-100 p-1 rounded-lg w-fit">
                            <button
                                onClick={() => setViewMode('week')}
                                className={`px-4 py-2 rounded-md font-medium transition-colors ${viewMode === 'week'
                                        ? 'bg-white text-blue-600 shadow-sm'
                                        : 'text-gray-600 hover:text-gray-800'
                                    }`}
                            >
                                Week View
                            </button>
                            <button
                                onClick={() => setViewMode('physicist')}
                                className={`px-4 py-2 rounded-md font-medium transition-colors ${viewMode === 'physicist'
                                        ? 'bg-white text-blue-600 shadow-sm'
                                        : 'text-gray-600 hover:text-gray-800'
                                    }`}
                            >
                                Physicist View
                            </button>
                        </div>
                    </div>

                    {viewMode === 'week' && (
                        <>
                            <div className="mb-4 border-b">
                                <div className="flex space-x-2 overflow-x-auto">
                                    {weeks.map((week, index) => (
                                        <button
                                            key={index}
                                            onClick={() => setCurrentWeek(index)}
                                            className={`px-3 py-2 font-semibold whitespace-nowrap flex-shrink-0 ${currentWeek === index
                                                    ? 'border-b-2 border-blue-500 text-blue-600'
                                                    : 'text-gray-600 hover:text-gray-800'
                                                }`}
                                        >
                                            Week {index + 1}
                                            <div className="text-xs">
                                                {week[0].toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - {week[4].toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
                                            </div>
                                        </button>
                                    ))}
                                </div>
                            </div>

                            <div className="overflow-x-auto">
                                <table className="w-full border-collapse border border-gray-300">
                                    <thead>
                                        <tr>
                                            <th className="border border-gray-300 p-2 bg-gray-100 w-48">Duty</th>
                                            {weeks[currentWeek].map((date, dayIndex) => (
                                                <th key={dayIndex} className="border border-gray-300 p-2 bg-gray-100 min-w-48">
                                                    {date.toLocaleDateString('en-US', { weekday: 'short' })}
                                                    <br />
                                                    {date.toLocaleDateString()}
                                                </th>
                                            ))}
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {duties.map((duty) => (
                                            <tr key={duty}>
                                                <td className="border border-gray-300 p-2 font-medium bg-gray-50">
                                                    <div className="flex items-center justify-between">
                                                        <span>{duty}</span>
                                                        <button
                                                            onClick={() => clearRow(duty)}
                                                            className="ml-2 p-1 text-red-600 hover:bg-red-100 rounded"
                                                            title="Clear entire row"
                                                        >
                                                            <Trash2 className="w-3 h-3" />
                                                        </button>
                                                    </div>
                                                </td>
                                                {weeks[currentWeek].map((date, dayIndex) => {
                                                    const dayOfWeek = date.getDay();
                                                    const currentDateStr = date.toISOString().split('T')[0];
                                                    const isRestricted = restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek);
                                                    const isHoliday = nationalHolidays.includes(currentDateStr);

                                                    if (isRestricted || isHoliday) {
                                                        return (
                                                            <td key={dayIndex} className="border border-gray-300 p-2 bg-gray-200 text-center">
                                                                {isHoliday ? 'HOLIDAY' : '-'}
                                                            </td>
                                                        );
                                                    }

                                                    const dateStr = date.toISOString().split('T')[0];
                                                    const currentAssignment = schedule[currentWeek]?.[dateStr]?.[duty] || '';
                                                    const { eligible, ineligible } = getEligiblePhysicists(duty, date, currentWeek);

                                                    const isCurrentAssignmentAvoidance = currentAssignment &&
                                                        availability[currentAssignment]?.soft?.includes(dateStr);

                                                    let hasConflicts = false;
                                                    if (currentAssignment) {
                                                        const currentAssignments = getCurrentAssignments(currentAssignment);
                                                        const currentCount = currentAssignments[duty] || 0;
                                                        let cap = dutyCaps[currentAssignment]?.[duty] || 0;
                                                        if (cap === 0 && combinedDutyMap[duty]) {
                                                            cap = dutyCaps[currentAssignment]?.[combinedDutyMap[duty]] || 0;
                                                        }

                                                        const isOverCap = currentCount > cap;

                                                        const daySchedule = schedule[currentWeek]?.[dateStr] || {};
                                                        const dayAssignments = Object.entries(daySchedule)
                                                            .filter(([d, p]) => p === currentAssignment && d !== duty)
                                                            .map(([d]) => d);

                                                        const isDoubleBooked = dayAssignments.length > 0 && !allowedCombinations.some(combo =>
                                                            combo.includes(duty) && dayAssignments.some(assigned => combo.includes(assigned))
                                                        );

                                                        hasConflicts = isOverCap || isDoubleBooked;
                                                    }

                                                    return (
                                                        <td
                                                            key={dayIndex}
                                                            className={`border border-gray-300 p-1 ${hasConflicts ? 'bg-blue-800 text-white' :
                                                                    isCurrentAssignmentAvoidance ? 'bg-red-100' :
                                                                        !currentAssignment ? 'bg-purple-50' : ''
                                                                }`}
                                                        >
                                                            <select
                                                                value={currentAssignment}
                                                                onChange={(e) => handleAssignment(duty, date, currentWeek, e.target.value)}
                                                                className={`w-full p-1 border rounded text-sm ${hasConflicts ? 'bg-blue-700 text-white' :
                                                                        isCurrentAssignmentAvoidance ? 'bg-red-50' :
                                                                            !currentAssignment ? 'bg-purple-50' : 'bg-white'
                                                                    }`}
                                                            >
                                                                <option value="">Select physicist...</option>
                                                                {eligible.map(physicist => {
                                                                    const isCurrentlyAssigned = currentAssignment === physicist.name;
                                                                    const showWarnings = isCurrentlyAssigned && hasConflicts;

                                                                    const wasOverAssignedLastWeek = wasOverAssignedPreviousWeek(physicist.name, duty, currentWeek);

                                                                    const daySchedule = schedule[currentWeek]?.[dateStr] || {};
                                                                    const isAssignedElsewhere = Object.entries(daySchedule).some(([otherDuty, assignedPhysicist]) =>
                                                                        assignedPhysicist === physicist.name && otherDuty !== duty
                                                                    );

                                                                    const warningText = showWarnings ? ` - ${physicist.warnings.join(', ')}` : '';
                                                                    let prefix = '';

                                                                    if (isAssignedElsewhere) {
                                                                        prefix = '✅ ';
                                                                    } else if (physicist.isAvoidance) {
                                                                        prefix = '🚨 ';
                                                                    } else if (showWarnings) {
                                                                        prefix = '⚠️ ';
                                                                    } else if (wasOverAssignedLastWeek) {
                                                                        prefix = '❗ ';
                                                                    }

                                                                    return (
                                                                        <option
                                                                            key={physicist.name}
                                                                            value={physicist.name}
                                                                            className={hasConflicts ? 'text-white' : ''}
                                                                        >
                                                                            {prefix}{physicist.name} ({physicist.currentCount}/{physicist.cap}) [{physicist.daysAvailable}/{physicist.avoidanceDays}]{warningText}
                                                                        </option>
                                                                    );
                                                                })}
                                                                {ineligible.length > 0 && (
                                                                    <optgroup label="--- Ineligible ---">
                                                                        {ineligible.map(physicist => (
                                                                            <option
                                                                                key={physicist.name}
                                                                                value={physicist.name}
                                                                                disabled
                                                                                className={hasConflicts ? "text-gray-300" : "text-gray-400"}
                                                                            >
                                                                                {physicist.name} - {physicist.reason}
                                                                            </option>
                                                                        ))}
                                                                    </optgroup>
                                                                )}
                                                            </select>
                                                        </td>
                                                    );
                                                })}
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </>
                    )}

                    {viewMode === 'physicist' && isDataReady && physicists.length > 0 && (
                        <>
                            <div className="mb-4 border-b">
                                <div className="flex justify-between items-center">
                                    <div className="flex space-x-2 overflow-x-auto">
                                        {physicists.map((physicist, index) => (
                                            <button
                                                key={index}
                                                onClick={() => setCurrentPhysicist(index)}
                                                className={`px-3 py-2 font-semibold whitespace-nowrap flex-shrink-0 ${currentPhysicist === index
                                                        ? 'border-b-2 border-blue-500 text-blue-600'
                                                        : 'text-gray-600 hover:text-gray-800'
                                                    }`}
                                            >
                                                {physicist}
                                                <div className="text-xs">
                                                    {(() => {
                                                        const assignments = getCurrentAssignments(physicist);
                                                        const totalAssignments = Object.values(assignments).reduce((sum, count) => sum + count, 0);
                                                        return `${totalAssignments} assignments`;
                                                    })()}
                                                </div>
                                            </button>
                                        ))}
                                    </div>

                                    <div className="flex-shrink-0 ml-4 flex space-x-2">
                                        <button
                                            onClick={() => autofillPhysicist(physicists[currentPhysicist])}
                                            className="bg-green-500 text-white px-3 py-2 rounded hover:bg-green-600 flex items-center text-sm"
                                            title={`Automatically fill remaining capacity for ${physicists[currentPhysicist]}`}
                                        >
                                            <CheckCircle className="w-4 h-4 mr-1" />
                                            Fill {physicists[currentPhysicist].split(' ')[0]}
                                        </button>
                                        <button
                                            onClick={() => clearPhysicistAllAvoidanceDays(physicists[currentPhysicist])}
                                            className="bg-orange-500 text-white px-3 py-2 rounded hover:bg-orange-600 flex items-center text-sm"
                                            title={`Clear all assignments on avoidance days for ${physicists[currentPhysicist]}`}
                                        >
                                            ⚠️
                                            Clear All Avoidance
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <div className="overflow-x-auto">
                                <table className="w-full border-collapse border border-gray-300">
                                    <thead>
                                        <tr>
                                            <th className="border border-gray-300 p-2 bg-gray-100 w-24">Week</th>
                                            <th className="border border-gray-300 p-2 bg-gray-100 min-w-48">Monday</th>
                                            <th className="border border-gray-300 p-2 bg-gray-100 min-w-48">Tuesday</th>
                                            <th className="border border-gray-300 p-2 bg-gray-100 min-w-48">Wednesday</th>
                                            <th className="border border-gray-300 p-2 bg-gray-100 min-w-48">Thursday</th>
                                            <th className="border border-gray-300 p-2 bg-gray-100 min-w-48">Friday</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {weeks.map((week, weekIndex) => (
                                            <tr key={weekIndex}>
                                                <td className="border border-gray-300 p-2 font-medium bg-gray-50 text-center">
                                                    <div className="flex flex-col items-center">
                                                        <div>Week {weekIndex + 1}</div>
                                                        <div className="text-xs mb-2">
                                                            {week[0].toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
                                                        </div>
                                                        <div className="flex flex-col space-y-1">
                                                            <button
                                                                onClick={() => clearPhysicistWeek(physicists[currentPhysicist], weekIndex)}
                                                                className="px-2 py-1 text-xs bg-red-500 text-white rounded hover:bg-red-600 flex items-center"
                                                                title={`Clear all assignments for ${physicists[currentPhysicist]} in Week ${weekIndex + 1}`}
                                                            >
                                                                <Trash2 className="w-3 h-3 mr-1" />
                                                                Clear Week
                                                            </button>
                                                            <button
                                                                onClick={() => clearPhysicistAvoidanceDays(physicists[currentPhysicist], weekIndex)}
                                                                className="px-2 py-1 text-xs bg-orange-500 text-white rounded hover:bg-orange-600 flex items-center"
                                                                title={`Clear assignments on avoidance days for ${physicists[currentPhysicist]} in Week ${weekIndex + 1}`}
                                                            >
                                                                ⚠️
                                                                Clear Avoidance
                                                            </button>
                                                        </div>
                                                    </div>
                                                </td>
                                                {week.map((date, dayIndex) => {
                                                    const dateStr = date.toISOString().split('T')[0];
                                                    const dayOfWeek = date.getDay();
                                                    const isHoliday = nationalHolidays.includes(dateStr);
                                                    const currentPhysicistName = physicists[currentPhysicist];

                                                    const daySchedule = schedule[weekIndex]?.[dateStr] || {};
                                                    const assignedDuties = Object.entries(daySchedule)
                                                        .filter(([duty, assignedPhysicist]) => assignedPhysicist === currentPhysicistName)
                                                        .map(([duty]) => duty);

                                                    const hasHardUnavailability = availability[currentPhysicistName]?.hard?.includes(dateStr) &&
                                                        !availability[currentPhysicistName]?.soft?.includes(dateStr);
                                                    const hasSoftAvoidance = availability[currentPhysicistName]?.soft?.includes(dateStr);

                                                    if (isHoliday) {
                                                        return (
                                                            <td key={dayIndex} className="border border-gray-300 p-2 bg-gray-200 text-center">
                                                                <div className="font-medium">HOLIDAY</div>
                                                                <div className="text-xs">{date.getDate()}</div>
                                                            </td>
                                                        );
                                                    }

                                                    if (hasHardUnavailability) {
                                                        return (
                                                            <td key={dayIndex} className="border border-gray-300 p-2 bg-red-200 text-center">
                                                                <div className="font-medium text-red-800">UNAVAILABLE</div>
                                                                <div className="text-xs">{date.getDate()}</div>
                                                            </td>
                                                        );
                                                    }

                                                    return (
                                                        <td
                                                            key={dayIndex}
                                                            className={`border border-gray-300 p-1 ${hasSoftAvoidance ? 'bg-yellow-100' :
                                                                    assignedDuties.length === 0 ? 'bg-purple-50' : ''
                                                                }`}
                                                        >
                                                            <div className="text-xs font-medium mb-1 text-center">
                                                                {date.getDate()}
                                                                {hasSoftAvoidance && <span className="text-yellow-600 ml-1">⚠️</span>}
                                                            </div>

                                                            {assignedDuties.map((duty, idx) => (
                                                                <div key={idx} className="mb-1 p-1 bg-blue-100 rounded text-xs">
                                                                    <div className="flex justify-between items-center">
                                                                        <span className="truncate flex-1" title={duty}>{duty}</span>
                                                                        <button
                                                                            onClick={() => handleAssignment(duty, date, weekIndex, '')}
                                                                            className="ml-1 text-red-600 hover:bg-red-200 rounded px-1"
                                                                            title="Remove assignment"
                                                                        >
                                                                            ×
                                                                        </button>
                                                                    </div>
                                                                </div>
                                                            ))}

                                                            <select
                                                                value=""
                                                                onChange={(e) => {
                                                                    if (e.target.value) {
                                                                        handlePhysicistAssignment(e.target.value, date, weekIndex, currentPhysicistName);
                                                                        e.target.value = '';
                                                                    }
                                                                }}
                                                                className="w-full p-1 border rounded text-xs bg-white"
                                                            >
                                                                <option value="">+ Add duty...</option>
                                                                {duties.map(duty => {
                                                                    const isRestricted = restrictedDuties[duty] && !restrictedDuties[duty].includes(dayOfWeek);
                                                                    if (isRestricted) return null;

                                                                    const alreadyAssigned = assignedDuties.includes(duty);
                                                                    if (alreadyAssigned) return null;

                                                                    const daySchedule = schedule[weekIndex]?.[dateStr] || {};
                                                                    const assignedToOther = daySchedule[duty] && daySchedule[duty] !== currentPhysicistName;

                                                                    let cap = dutyCaps[currentPhysicistName]?.[duty] || 0;
                                                                    if (cap === 0 && combinedDutyMap[duty]) {
                                                                        cap = dutyCaps[currentPhysicistName]?.[combinedDutyMap[duty]] || 0;
                                                                    }

                                                                    if (cap === 0) return null;

                                                                    const currentAssignments = getCurrentAssignments(currentPhysicistName);
                                                                    const currentCount = currentAssignments[duty] || 0;
                                                                    const isOverCap = currentCount >= cap;

                                                                    let displayText = duty;
                                                                    let className = '';

                                                                    if (assignedToOther) {
                                                                        displayText = `⚠️ ${duty} (assigned to ${daySchedule[duty]})`;
                                                                        className = 'text-orange-600 font-medium';
                                                                    } else if (isOverCap) {
                                                                        displayText = `${duty} (${currentCount}/${cap} - Over!)`;
                                                                        className = 'text-red-600';
                                                                    } else {
                                                                        displayText = `${duty} (${currentCount}/${cap})`;
                                                                    }

                                                                    return (
                                                                        <option
                                                                            key={duty}
                                                                            value={duty}
                                                                            className={className}
                                                                        >
                                                                            {displayText}
                                                                        </option>
                                                                    );
                                                                })}
                                                            </select>
                                                        </td>
                                                    );
                                                })}
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            <div className="mt-4 p-4 bg-gray-50 rounded-lg">
                                <h3 className="font-semibold mb-2">{physicists[currentPhysicist]} - Assignment Summary</h3>
                                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                                    {(() => {
                                        const assignments = getCurrentAssignments(physicists[currentPhysicist]);
                                        return duties.map(duty => {
                                            const count = assignments[duty] || 0;
                                            let cap = dutyCaps[physicists[currentPhysicist]]?.[duty] || 0;
                                            if (cap === 0 && combinedDutyMap[duty]) {
                                                cap = dutyCaps[physicists[currentPhysicist]]?.[combinedDutyMap[duty]] || 0;
                                            }

                                            if (cap === 0 && count === 0) return null;

                                            const isOverCap = count > cap;
                                            const utilization = cap > 0 ? (count / cap * 100).toFixed(0) : 'N/A';

                                            return (
                                                <div key={duty} className={`p-2 rounded border ${isOverCap ? 'bg-red-100 border-red-300' : 'bg-white border-gray-200'}`}>
                                                    <div className="font-medium text-sm">{duty}</div>
                                                    <div className={`text-lg font-bold ${isOverCap ? 'text-red-600' : 'text-blue-600'}`}>
                                                        {count}/{cap}
                                                    </div>
                                                    <div className="text-xs text-gray-600">
                                                        {cap > 0 ? `${utilization}% utilized` : 'No capacity'}
                                                    </div>
                                                </div>
                                            );
                                        }).filter(Boolean);
                                    })()}
                                </div>
                            </div>
                        </>
                    )}
                </>
            )}

            {showConflictModal && conflictData && (
                <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                    <div className="bg-white p-6 rounded-lg shadow-xl max-w-md w-full mx-4">
                        <h3 className="text-lg font-bold mb-4 text-red-600">Scheduling Conflict Detected</h3>
                        <div className="mb-4">
                            <p className="mb-2">
                                <strong>{conflictData.physicist}</strong> is already assigned to{' '}
                                <strong>"{conflictData.existingDuty}"</strong> on{' '}
                                <strong>{conflictData.date.toLocaleDateString()}</strong>.
                            </p>
                            <p className="text-sm text-gray-600 mb-4">
                                You are trying to assign them to <strong>"{conflictData.newDuty}"</strong> on the same date.
                            </p>
                            <p className="text-sm text-gray-700">What would you like to do?</p>
                        </div>

                        <div className="flex flex-col space-y-3">
                            <button
                                onClick={() => handleConflictResolution(true)}
                                className="bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600 font-medium"
                            >
                                🔄 Switch Assignment
                                <div className="text-xs mt-1 opacity-90">
                                    Remove from "{conflictData.existingDuty}" and assign to "{conflictData.newDuty}"
                                </div>
                            </button>

                            <button
                                onClick={() => handleConflictResolution(false)}
                                className="bg-orange-500 text-white px-4 py-2 rounded hover:bg-orange-600 font-medium"
                            >
                                ⚠️ Double Book
                                <div className="text-xs mt-1 opacity-90">
                                    Keep "{conflictData.existingDuty}" and also assign to "{conflictData.newDuty}"
                                </div>
                            </button>

                            <button
                                onClick={() => {
                                    setShowConflictModal(false);
                                    setConflictData(null);
                                }}
                                className="bg-gray-500 text-white px-4 py-2 rounded hover:bg-gray-600 font-medium"
                            >
                                ❌ Cancel
                                <div className="text-xs mt-1 opacity-90">
                                    Don't make any changes
                                </div>
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {showDutyConflictModal && dutyConflictData && (
                <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                    <div className="bg-white p-6 rounded-lg shadow-xl max-w-md w-full mx-4">
                        <h3 className="text-lg font-bold mb-4 text-orange-600 flex items-center">
                            ⚠️ Duty Already Assigned
                        </h3>
                        <div className="mb-4">
                            <p className="mb-2">
                                <strong>"{dutyConflictData.duty}"</strong> on{' '}
                                <strong>{dutyConflictData.date.toLocaleDateString()}</strong> is already assigned to{' '}
                                <strong>{dutyConflictData.currentlyAssigned}</strong>.
                            </p>
                            <p className="text-sm text-gray-600 mb-4">
                                You are trying to assign it to <strong>{dutyConflictData.newPhysicist}</strong>.
                            </p>
                            <p className="text-sm text-gray-700">Would you like to reassign this duty?</p>
                        </div>

                        <div className="flex flex-col space-y-3">
                            <button
                                onClick={() => handleDutyConflictResolution(true)}
                                className="bg-orange-500 text-white px-4 py-2 rounded hover:bg-orange-600 font-medium"
                            >
                                🔄 Reassign Duty
                                <div className="text-xs mt-1 opacity-90">
                                    Move from "{dutyConflictData.currentlyAssigned}" to "{dutyConflictData.newPhysicist}"
                                </div>
                            </button>

                            <button
                                onClick={() => handleDutyConflictResolution(false)}
                                className="bg-gray-500 text-white px-4 py-2 rounded hover:bg-gray-600 font-medium"
                            >
                                ❌ Cancel
                                <div className="text-xs mt-1 opacity-90">
                                    Keep current assignment
                                </div>
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {showExportModal && (
                <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                    <div className="bg-white p-6 rounded-lg shadow-xl max-w-md w-full mx-4">
                        <h3 className="text-lg font-bold mb-4 text-green-600 flex items-center">
                            <Download className="w-5 h-5 mr-2" />
                            Excel File Ready
                        </h3>
                        <div className="text-center p-6">
                            <div className="text-6xl mb-4">📊</div>
                            <p className="text-lg font-semibold mb-2">Your physicist schedule is ready!</p>
                            <p className="text-sm text-gray-600 mb-6">
                                Click below to download the Excel file with proper formatting.
                            </p>
                            <button
                                onClick={() => {
                                    if (weeks && weeks.length > 0) {
                                        const endDate = weeks[weeks.length - 1][4].toISOString().split('T')[0];
                                        const filename = `physicist_schedule_${startDate.replace(/-/g, '')}_to_${endDate.replace(/-/g, '')}.xlsx`;

                                        const a = document.createElement('a');
                                        a.href = exportContent;
                                        a.download = filename;
                                        a.style.display = 'none';
                                        document.body.appendChild(a);
                                        a.click();
                                        document.body.removeChild(a);

                                        setTimeout(() => {
                                            setShowExportModal(false);
                                            setExportContent('');
                                        }, 500);
                                    }
                                }}
                                className="inline-flex items-center bg-green-500 text-white px-6 py-3 rounded-lg hover:bg-green-600 font-medium text-lg mb-4"
                            >
                                <Download className="w-5 h-5 mr-2" />
                                Download Excel File
                            </button>
                            <div className="text-xs text-gray-500">
                                Excel format • {duties.length} duties • {weeks.length} weeks
                            </div>
                        </div>

                        <div className="flex justify-center">
                            <button
                                onClick={() => {
                                    setShowExportModal(false);
                                    setExportContent('');
                                }}
                                className="bg-gray-500 text-white px-4 py-2 rounded hover:bg-gray-600 font-medium"
                            >
                                Close
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

function App() {
    return (
        <div className="App">
            <PhysicistScheduler />
        </div>
    );
}

export default App;