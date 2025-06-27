// Global variables
let scheduleData = null;
let processedEvents = [];
// No global activityReminderCheckboxes; always get by ID when needed
let activityReminderMap = {};

// DOM elements
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const enableReminders = document.getElementById('enableReminders');
const reminderSettings = document.getElementById('reminderSettings');
const reminderTime = document.getElementById('reminderTime');
const previewSection = document.getElementById('previewSection');
const previewBtn = document.getElementById('previewBtn');
const previewTable = document.getElementById('previewTable');
const eventCount = document.getElementById('eventCount');
const exportBtn = document.getElementById('exportBtn');
const exportStatus = document.getElementById('exportStatus');

// Event listeners
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM fully loaded');
    fileInput.addEventListener('change', function(e) {
        console.log('File input changed', e.target.files);
    });
    fileInput.addEventListener('change', handleFileUpload);
    enableReminders.addEventListener('change', toggleReminderSettings);
    previewBtn.addEventListener('click', generatePreview);
    exportBtn.addEventListener('click', exportCalendar);
    
    // Initialize reminder settings visibility
    toggleReminderSettings();
});

// File upload handler
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON with headers
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length < 2) {
                showError('Invalid file format. Please ensure the file has headers and data.');
                return;
            }

            // Extract headers and data
            const headers = jsonData[0];
            const rows = jsonData.slice(1);
            
            // Parse the data
            scheduleData = parseScheduleData(headers, rows);
            
            // Show file info
            showFileInfo(file, scheduleData.length);
            
            // Enable export button
            exportBtn.disabled = false;
            
            // Show preview section
            previewSection.classList.remove('hidden');
            // Always generate per-activity reminder checkboxes after file upload
            showActivityReminderCheckboxes();
            // Show/hide checkboxes based on reminder toggle
            toggleReminderSettings();
        } catch (error) {
            console.error('Error parsing file:', error);
            showError('Error parsing the Excel file. Please ensure it\'s a valid .xlsx file.');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Parse schedule data from Excel
function parseScheduleData(headers, rows) {
    // Normalize headers: trim and lowercase
    const normHeaders = headers.map(h => h ? h.toString().trim().toLowerCase() : '');
    console.log('Normalized headers:', normHeaders);

    // Find column indices (case-insensitive)
    const columnMap = {
        from: normHeaders.findIndex(h => h === 'from'),
        to: normHeaders.findIndex(h => h === 'to'),
        version: normHeaders.findIndex(h => h === 'version'),
        bkd: normHeaders.findIndex(h => h === 'bkd'),
        // Support both 'private activity' and 'activity' as possible headers
        activity: normHeaders.findIndex(h => h === 'private activity') !== -1 ? normHeaders.findIndex(h => h === 'private activity') : normHeaders.findIndex(h => h === 'activity'),
        unit: normHeaders.findIndex(h => h === 'unit'),
        fromLoc: normHeaders.findIndex(h => h === 'from loc'),
        toLoc: normHeaders.findIndex(h => h === 'to loc'),
        area: normHeaders.findIndex(h => h === 'area'),
        notes: normHeaders.findIndex(h => h === 'notes'),
        departureAttributes: normHeaders.findIndex(h => h === 'departure attributes')
    };
    console.log('Detected column indices:', columnMap);

    // Validate required columns
    const requiredColumns = ['from', 'to', 'version', 'bkd', 'activity'];
    const missingColumns = requiredColumns.filter(col => columnMap[col] === -1);
    if (missingColumns.length > 0) {
        console.error('Missing required columns:', missingColumns);
        throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
    }

    let lastArea = null;
    let firstArea = null;
    const parsedData = [];
    const rowsToBackfill = [];

    rows.forEach((row, index) => {
        // Skip empty rows
        if (!row[columnMap.from] && !row[columnMap.activity]) {
            console.log(`Row ${index + 2} skipped: missing 'From' and 'Activity'`);
            return;
        }

        const data = {
            from: row[columnMap.from],
            to: row[columnMap.to] || row[columnMap.from],
            version: row[columnMap.version] || '',
            bkd: row[columnMap.bkd] || '',
            activity: row[columnMap.activity] || '',
            unit: row[columnMap.unit] || '',
            fromLoc: row[columnMap.fromLoc] || '',
            toLoc: row[columnMap.toLoc] || '',
            area: row[columnMap.area] || '',
            notes: row[columnMap.notes] || '',
            departureAttributes: row[columnMap.departureAttributes] || '',
            rowIndex: index + 2 // +2 because we start from row 2 (after headers)
        };

        // Handle area inheritance and backfill
        if (data.area) {
            lastArea = data.area;
            if (!firstArea) {
                firstArea = data.area;
                // Backfill all previous rows with empty area
                rowsToBackfill.forEach(prevData => {
                    prevData.area = firstArea;
                });
                rowsToBackfill.length = 0;
            }
        } else if (lastArea) {
            data.area = lastArea;
        } else if (firstArea) {
            data.area = firstArea;
        } else {
            // Save for backfilling once firstArea is found
            rowsToBackfill.push(data);
        }

        console.log(`Parsed row ${index + 2}:`, data);
        parsedData.push(data);
    });

    // Final backfill if firstArea is found after parsing
    if (firstArea) {
        rowsToBackfill.forEach(prevData => {
            prevData.area = firstArea;
        });
    }

    console.log('Total parsed rows:', parsedData.length);
    return parsedData;
}

// Update parseDate to use Luxon without timezone
function parseDate(dateVal) {
    if (!dateVal) return null;
    if (typeof dateVal === 'number') {
        // Excel serial date to JS Date (local time, not UTC)
        const utc_days = dateVal - 25569;
        const utc_value = utc_days * 86400 * 1000;
        const date = new Date(utc_value);
        // Use Luxon to get correct local date
        return luxon.DateTime.fromObject({
            year: date.getUTCFullYear(),
            month: date.getUTCMonth() + 1,
            day: date.getUTCDate()
        });
    }
    // Otherwise, try to parse as mm/dd/yy
    const parts = dateVal.toString().split('/');
    if (parts.length !== 3) return null;
    const month = parseInt(parts[0]);
    const day = parseInt(parts[1]);
    const year = parseInt(parts[2]) + 2000;
    return luxon.DateTime.fromObject({ year, month, day });
}

// Generate calendar events
function generateCalendarEvents() {
    if (!scheduleData) return [];
    
    const timeSettings = getTimeSettings();
    const reminderSettings = getReminderSettings();
    
    return scheduleData.map(row => {
        const fromDate = parseDate(row.from);
        const toDate = parseDate(row.to);
        
        if (!fromDate || !toDate) {
            console.warn(`Skipping row ${row.rowIndex}: invalid dates`);
            return null;
        }
        
        return createEvent(row, fromDate, toDate, timeSettings, reminderSettings);
    }).filter(Boolean);
}

// Update createEvent to use Luxon for reminder calculation
function createEvent(row, fromDate, toDate, timeSettings, reminderSettings) {
    const subject = `${row.activity} - ${row.version}`.trim();
    
    // Set location based on activity type
    let location;
    if (row.activity === 'Drive Unit' || row.activity === 'Unload Unit') {
        location = (row.toLoc || '').replace(/_/g, ' ');
    } else {
        location = (row.area || '').replace(/_/g, ' ');
    }
    
    let description = `${row.activity} ${row.version}`;
    
    if (row.activity === 'Travel') {
        description = `Travel from ${row.fromLoc || 'Unknown'} to ${row.toLoc || 'Unknown'}`;
    } else if (row.activity === 'Prep Only' || row.activity === 'Prep Trip') {
        description += `\nGuests: ${row.bkd}`;
        description += `\nUnit: ${row.unit}`;
        if (row.departureAttributes) {
            description += `\nNotes: ${row.departureAttributes}`;
        }
    }
    let startTime, endTime, allDay = false;
    switch (row.activity) {
        case 'Prep Trip':
            startTime = timeSettings.prepTripStart;
            endTime = timeSettings.prepTripEnd;
            break;
        case 'Prep Only':
            startTime = timeSettings.prepOnlyStart;
            endTime = timeSettings.prepOnlyEnd;
            break;
        case 'Trip Breakdown':
            startTime = timeSettings.breakdownStart;
            endTime = timeSettings.breakdownEnd;
            break;
        case 'Drive Unit':
            allDay = true;
            break;
        case 'Load Unit':
            startTime = '09:00';
            endTime = '15:00';
            break;
        default:
            startTime = '09:00';
            endTime = '17:00';
    }
    // Use Luxon for event start/end
    let eventStart = fromDate;
    let eventEnd = toDate;
    if (!allDay && startTime) {
        const [h, m] = startTime.split(':');
        eventStart = eventStart.set({ hour: parseInt(h), minute: parseInt(m), second: 0, millisecond: 0 });
    } else {
        eventStart = eventStart.set({ hour: 0, minute: 0, second: 0, millisecond: 0 });
    }
    if (!allDay && endTime) {
        const [h, m] = endTime.split(':');
        eventEnd = eventEnd.set({ hour: parseInt(h), minute: parseInt(m), second: 0, millisecond: 0 });
    } else {
        eventEnd = eventEnd.set({ hour: 23, minute: 59, second: 59, millisecond: 999 });
    }
    const event = {
        subject: subject,
        description: description,
        location: location,
        fromDate: eventStart,
        toDate: eventEnd,
        startTime: startTime,
        endTime: endTime,
        allDay: allDay,
        hasTravel: !!(row.fromLoc || row.toLoc),
        fromLoc: row.fromLoc,
        toLoc: row.toLoc
    };
    if (reminderSettings.enabled && reminderSettings.activityMap[row.activity]) {
        let reminderDate;
        if (allDay) {
            // For all-day events, calculate reminder based on 8:00 AM start time
            const eventStartAt8AM = eventStart.set({ hour: 8, minute: 0, second: 0, millisecond: 0 });
            reminderDate = eventStartAt8AM.minus({ hours: reminderSettings.hours });
        } else {
            // For timed events, use the actual start time
            reminderDate = eventStart.minus({ hours: reminderSettings.hours });
        }
        event.reminderDate = reminderDate;
    }
    return event;
}

// Get time settings from form
function getTimeSettings() {
    return {
        prepTripStart: document.getElementById('prepTripStart').value,
        prepTripEnd: document.getElementById('prepTripEnd').value,
        breakdownStart: document.getElementById('breakdownStart').value,
        breakdownEnd: document.getElementById('breakdownEnd').value,
        prepOnlyStart: document.getElementById('prepOnlyStart').value,
        prepOnlyEnd: document.getElementById('prepOnlyEnd').value
    };
}

// Get reminder settings from form
function getReminderSettings() {
    return {
        enabled: enableReminders.checked,
        hours: parseFloat(reminderTime.value) || 24,
        activityMap: { ...activityReminderMap }
    };
}

// Toggle reminder settings visibility
function toggleReminderSettings() {
    console.log('toggleReminderSettings called, enableReminders.checked:', enableReminders.checked);
    const container = document.getElementById('activityReminderCheckboxes');
    console.log('activityReminderCheckboxes container:', container);
    console.log('scheduleData exists:', !!scheduleData);
    
    if (enableReminders.checked) {
        reminderSettings.classList.remove('hidden');
        // Only show activity checkboxes if we have schedule data (file uploaded)
        if (container && scheduleData) {
            console.log('Showing activity checkboxes');
            container.style.display = 'flex';
        } else {
            console.log('Not showing checkboxes - container:', !!container, 'scheduleData:', !!scheduleData);
        }
    } else {
        reminderSettings.classList.add('hidden');
        if (container) {
            console.log('Hiding activity checkboxes');
            container.style.display = 'none';
        }
    }
}

// Show file info
function showFileInfo(file, rowCount) {
    fileInfo.innerHTML = `
        <strong>File loaded:</strong> ${file.name}<br>
        <strong>Size:</strong> ${(file.size / 1024).toFixed(1)} KB<br>
        <strong>Rows processed:</strong> ${rowCount}
    `;
    fileInfo.classList.remove('hidden');
}

// Generate preview
function generatePreview() {
    if (!scheduleData) {
        showError('Please upload a file first.');
        return;
    }

    processedEvents = generateCalendarEvents();
    
    if (processedEvents.length === 0) {
        showError('No valid events found in the schedule.');
        return;
    }

    displayPreview(processedEvents);
    eventCount.textContent = `${processedEvents.length} events generated`;
}

// Display preview table
function displayPreview(events) {
    const format = document.querySelector('input[name="exportFormat"]:checked').value;
    const remindersEnabled = getReminderSettings().enabled;
    let headers, rows;

    if (format === 'google') {
        headers = ['Subject', 'Start Date', 'End Date', 'Start Time', 'End Time', 'All Day Event'];
        if (remindersEnabled) {
            headers.push('Reminder Date', 'Reminder Time');
        }
        headers.push('Description', 'Location');
        rows = events.map(event => {
            const rowArr = [
                event.subject,
                formatDate(event.fromDate),
                formatDate(event.toDate),
                event.startTime,
                event.endTime,
                event.allDay ? 'True' : 'False'
            ];
            if (remindersEnabled) {
                rowArr.push(
                    event.reminderDate ? formatDate(event.reminderDate) : '',
                    event.reminderDate ? formatTime(event.reminderDate) : ''
                );
            }
            rowArr.push(event.description, event.location);
            return rowArr;
        });
    } else {
        headers = ['Subject', 'Start Time', 'End Time', 'All day event', 'Reminder Date', 'Reminder Time', 'Description', 'Location'];
        rows = events.map(event => [
            event.subject,
            formatDateTime(event.fromDate, event.startTime),
            formatDateTime(event.toDate, event.endTime),
            event.allDay ? 'True' : 'False',
            event.reminderDate ? formatDate(event.reminderDate) : '',
            event.reminderDate ? formatTime(event.reminderDate) : '',
            event.description,
            event.location
        ]);
    }

    const table = createTable(headers, rows);
    previewTable.innerHTML = '';
    previewTable.appendChild(table);
}

// Create table element
function createTable(headers, rows) {
    const table = document.createElement('table');
    
    // Create header row
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body rows
    const tbody = document.createElement('tbody');
    rows.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    
    return table;
}

// Export calendar
function exportCalendar() {
    if (!processedEvents || processedEvents.length === 0) {
        showError('Please generate a preview first.');
        return;
    }

    const format = document.querySelector('input[name="exportFormat"]:checked').value;
    
    try {
        if (format === 'google') {
            exportGoogleCalendar(processedEvents);
        } else {
            exportOutlookCalendar(processedEvents);
        }
        
        showSuccess(`Calendar exported successfully as ${format === 'google' ? 'CSV' : 'ICS'} file!`);
    } catch (error) {
        console.error('Export error:', error);
        showError('Error exporting calendar. Please try again.');
    }
}

// Export Google Calendar CSV
function exportGoogleCalendar(events) {
    const headers = ['Subject', 'Start Date', 'End Date', 'Start Time', 'End Time', 'All Day Event', 'Reminder Date', 'Reminder Time', 'Description', 'Location'];
    
    const csvContent = [
        headers.join(','),
        ...events.map(event => {
            let reminderDate = '';
            let reminderTime = '';
            if (event.reminderDate) {
                reminderDate = formatDate(event.reminderDate);
                reminderTime = formatTime(event.reminderDate);
            }
            return [
                `"${event.subject}"`,
                formatDate(event.fromDate),
                formatDate(event.toDate),
                event.startTime,
                event.endTime,
                event.allDay ? 'True' : 'False',
                reminderDate,
                reminderTime,
                `"${event.description.replace(/"/g, '""')}"`,
                `"${event.location}"`
            ].join(',');
        })
    ].join('\n');

    downloadFile(csvContent, 'calendar_events.csv', 'text/csv');
}

// Export Outlook Calendar ICS
function exportOutlookCalendar(events) {
    const icsEvents = events.map(event => {
        // Format datetime for ICS (local timezone)
        let startDateTime, endDateTime;
        
        if (event.allDay) {
            startDateTime = formatDate(event.fromDate);
            endDateTime = formatDate(event.toDate);
        } else {
            // Use local timezone formatting
            startDateTime = formatDateTime(event.fromDate, event.startTime);
            endDateTime = formatDateTime(event.toDate, event.endTime);
        }

        let icsEvent = [
            'BEGIN:VEVENT',
            `SUMMARY:${event.subject}`,
            `DTSTART${event.allDay ? ';VALUE=DATE' : ''}:${startDateTime}`,
            `DTEND${event.allDay ? ';VALUE=DATE' : ''}:${endDateTime}`,
            `DESCRIPTION:${event.description.replace(/\n/g, '\\n')}`,
            `LOCATION:${event.location}`,
        ];

        if (event.reminderDate) {
            icsEvent.push(`BEGIN:VALARM`);
            icsEvent.push(`TRIGGER:-PT${Math.floor(getReminderSettings().hours)}H`);
            icsEvent.push(`ACTION:DISPLAY`);
            icsEvent.push(`DESCRIPTION:Reminder`);
            icsEvent.push(`END:VALARM`);
        }

        icsEvent.push('END:VEVENT');
        return icsEvent.join('\r\n');
    });

    const icsContent = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Work Schedule Exporter//EN',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        ...icsEvents,
        'END:VCALENDAR'
    ].join('\r\n');

    downloadFile(icsContent, 'calendar_events.ics', 'text/calendar');
}

// Download file
function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Update formatting functions to use Luxon
function formatDate(date) {
    if (!date) return '';
    if (date.toFormat) return date.toFormat('yyyy-LL-dd');
    return date.toISOString().split('T')[0];
}
function formatTime(date) {
    if (!date) return '';
    if (date.toFormat) return date.toFormat('HH:mm');
    return date.toTimeString().split(' ')[0];
}
function formatDateTime(date, time) {
    if (!date) return '';
    if (date.toFormat) {
        if (time) {
            const [h, m] = time.split(':');
            return date.set({ hour: parseInt(h), minute: parseInt(m), second: 0, millisecond: 0 }).toFormat("yyyyLLdd'T'HHmmss");
        }
        return date.toFormat("yyyyLLdd'T'HHmmss");
    }
    const dateStr = formatDate(date);
    return `${dateStr}T${time}:00`;
}

function showError(message) {
    exportStatus.textContent = message;
    exportStatus.className = 'export-status error';
}

function showSuccess(message) {
    exportStatus.textContent = message;
    exportStatus.className = 'export-status success';
}

// Show per-activity reminder checkboxes after parsing
function showActivityReminderCheckboxes() {
    const container = document.getElementById('activityReminderCheckboxes');
    if (!container) {
        console.warn('activityReminderCheckboxes container is null');
        return;
    }
    if (!scheduleData) return;
    // Get unique activities
    const uniqueActivities = Array.from(new Set(scheduleData.map(row => row.activity).filter(Boolean)));
    // Preserve previous selections
    const previousMap = { ...activityReminderMap };
    container.innerHTML = '';
    activityReminderMap = {};
    uniqueActivities.forEach(activity => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        // Preserve previous selection, default to true if not set
        checkbox.checked = previousMap.hasOwnProperty(activity) ? previousMap[activity] : true;
        checkbox.value = activity;
        activityReminderMap[activity] = checkbox.checked;
        checkbox.addEventListener('change', (e) => {
            activityReminderMap[activity] = e.target.checked;
        });
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(activity));
        container.appendChild(label);
    });
    // Show if reminders are enabled, hide if disabled
    container.style.display = enableReminders.checked ? 'flex' : 'none';
} 