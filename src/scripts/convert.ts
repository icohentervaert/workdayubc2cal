import * as XLSX from "xlsx";

async function convert(file_name: string, file_url:string ) {
        const nameOfXLSX = file_name;
    
        const res = await fetch(file_url);
        const arrayBuffer = await res.arrayBuffer();
    
        // Read workbook
        const wb = XLSX.read(arrayBuffer, { type: "array" });
        const firstSheetName = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheetName];
    
        // AI generated code to make the library ensure that the proper headers are found 
        // and it doesn't just return nothing for all cells
        if (ws["!ref"] === "A1" || !ws["!ref"]) {
            const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
            Object.keys(ws)
                .filter((x) => x.charAt(0) !== "!")
                .forEach((x) => {
                    const res = XLSX.utils.decode_cell(x);
                    if (range.s.r > res.r) range.s.r = res.r;
                    if (range.s.c > res.c) range.s.c = res.c;
                    if (range.e.r < res.r) range.e.r = res.r;
                    if (range.e.c < res.c) range.e.c = res.c;
                });
            ws["!ref"] = XLSX.utils.encode_range(range);
        }
        
        // Convert worksheet to array-of-arrays to locate header row
        const rows: string[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
    
        let header_row_index: number | null = null;
        for (let i = 0; i < rows.length; i++) {
            if (rows[i]?.includes("Meeting Patterns")) {
                header_row_index = i;
                console.log(
                    `Found header 'Meeting Patterns' in sheet at row index ${header_row_index}`
                );
                break;
            }
        }
    
        if (header_row_index == null) {
            console.log("unable to find header!");
            throw new Error("RuntimeError");
        }
    
        // Build objects using detected header row
        const headers = rows[header_row_index];
        const dataRows = rows.slice(header_row_index + 1);
    
        const meetingPatternIndex = headers.indexOf("Meeting Patterns");
        const instructionFormatIndex = headers.indexOf("Instructional Format");
        const courseNameIndex = headers.findIndex(
            (h) => typeof h === "string" && h.includes("Course")
        );
    
        if (
            meetingPatternIndex === -1 ||
            instructionFormatIndex === -1 ||
            courseNameIndex === -1
        ) {
            throw new Error("No header containing found");
        }
    
        const courses = [];
    
        for (const r of dataRows) {
            if (!r || r.length === 0) continue;
    
            let name = r[courseNameIndex] ?? "";
            name = String(name);
    
            name = name.replace(/\s?\((Lecture|Discussion|Seminar|Laboratory)\)/gi, "");
            name = name.replace(/_V|_O/gi, "");
    
            const course = {
                name,
                type: r[instructionFormatIndex] ?? "",
                meeting_schedule: r[meetingPatternIndex] || "Online",
            };
    
            courses.push(course);
        }
    
        const events = [];
    
        for (const course of courses) {
            const meetingSchedule = course.meeting_schedule;
            if (meetingSchedule === "Online") {
                console.log(
                    `${course.name} has an online meeting schedule that is online, so please plan it yourself, as it will be skipped by the tool`
                );
                continue;
            }
    
            const name = `${course.name} (${course.type})`;
            const multipleSchedules = String(meetingSchedule).split("\n");
    
            for (let schedule of multipleSchedules) {
                if (!schedule.trim()) continue;
    
                let splitArr = schedule.split(" | ");
                console.log(splitArr);
    
                if (splitArr.length < 4) {
                    splitArr[2] = splitArr[2].replace(" |", "");
                    splitArr.push("Online");
                }
    
                console.log(splitArr.length);
    
                let startAndEndDate, days, times, location;
                if (splitArr.length === 7) {
                    const [a, b, c, campus, building, floor, room] = splitArr;
                    startAndEndDate = a;
                    days = b;
                    times = c;
                    location = `${building} ${floor} ${room}`;
                    console.log(location);
                } else {
                    [startAndEndDate, days, times, location] = splitArr;
                }
    
                const [startDate, endDate] = startAndEndDate.split(" - ");
                const [startTime, endTime] = times.split(" - ");
                const dayTokens = days.split(" ");
                const recurringDays = dayTokens.map((x) => x.slice(0, 2).toUpperCase());
    
                const beginStart = fixWrongStartDate(
                    meetingScheduleTime(startDate + startTime),
                    recurringDays
                );
                const beginEnd = fixWrongStartDate(
                    meetingScheduleTime(startDate + endTime),
                    recurringDays
                );
    
                const finalDay = parseDateOnly(endDate); // equivalent to DateTimeObj.strptime(endDate, "%Y-%m-%d")
    
                events.push(
                    createEvent(
                        name,
                        location,
                        [beginStart, beginEnd],
                        finalDay,
                        recurringDays
                    )
                );
            }
        }
    
        const icsText = serializeCalendar(events);
        downloadCalendar(icsText, nameOfXLSX.replace(".xlsx", ".ics"));

}

function generateUID() {
	// Generates a random string to act as a unique ID
	return (
		Math.random().toString(36).substring(2, 15) +
		Math.random().toString(36).substring(2, 15) +
		"@ubc_calendar_tool"
	);
}

function pad(n : number) {
	return String(n).padStart(2, "0");
}


function formatDateTimeLocal(dt: Date) {
	return `${dt.getFullYear()}${pad(dt.getMonth() + 1)}${pad(
		dt.getDate()
	)}T${pad(dt.getHours())}${pad(dt.getMinutes())}${pad(dt.getSeconds())}`;
}


function parseDateOnly(s: string) {
	const [y, m, d] = s.split("-").map(Number);
	return new Date(y, m - 1, d, 0, 0, 0);
}

function meetingScheduleTime(s: string) {
	// Python: time.replace(".", ""), "%Y-%m-%d%I:%M %p"
	const cleaned = s.replace(/\./g, ""); // remove dots
	const m = cleaned.match(
		/^(\d{4})-(\d{2})-(\d{2})(\d{1,2}):(\d{2})\s*(AM|PM)$/i
	);
	if (!m) throw new Error(`Invalid meeting time format: ${s}`);

	let [, Y, M, D, h, min, ampm]: [number,number,number,number,number,number,string] = m as any;
	Y = +Y;
	M = +M;
	D = +D;
	h = +h;
	min = +min;
	ampm = ampm.toUpperCase();

	if (ampm === "AM" && h === 12) h = 0;
	if (ampm === "PM" && h !== 12) h += 12;

	return new Date(Y, M - 1, D, h, min, 0);
}


function fixWrongStartDate(current_date: Date, recurring_days: string[]) {
	// Python mapping: {'SU': 6, 'MO': 0, ... 'SA': 5}
	// JS getDay(): SU=0..SA=6, convert to Python weekday() style MO=0..SU=6
	const days_map: Record<string, number> = { SU: 6, MO: 0, TU: 1, WE: 2, TH: 3, FR: 4, SA: 5 };
	const jsDay = current_date.getDay();
	const current_weekday = (jsDay + 6) % 7; // JS->Python weekday index

	const days_ahead = recurring_days.map(
		(day: string) => (days_map[day] - current_weekday + 7) % 7
	);
	const min_days_ahead = days_ahead.length ? Math.min(...days_ahead) : 0;

	const next = new Date(current_date);
	next.setDate(next.getDate() + min_days_ahead);
	return next;
}

function timeZoneOffset(timeZone: string, targetDate: Date) {
  const dt = new Date(Date.UTC(
    targetDate.getUTCFullYear(),
    targetDate.getUTCMonth(),
    targetDate.getUTCDate(),
    targetDate.getUTCHours(),
    targetDate.getUTCMinutes(),
    targetDate.getUTCSeconds()
  ));
  const local = new Date(dt.toLocaleString("en-US", { timeZone }));
  // @ts-ignore
  const mins = Math.round((local - dt) / 60000);

  const sign = mins >= 0 ? "+" : "-";
  const abs = Math.abs(mins);
  const hh = pad(Math.floor(abs / 60));
  const mm = pad(abs % 60);
  return `${sign}${hh}${mm}`;
}

function createEvent(
	name: string,
	location: string,
	beginDay: [Date,Date],
	finalDay: Date,
	recurringDays: string[],
	timeZone = "America/Vancouver"
) {
	const [dtStart, dtEnd] = beginDay;

	// Keep the odd Python behavior:
	// int(timeZoneOffset(timeZone).replace("0", ""))*-1 - 1
	// (not recommended, but preserved)
	const tzRaw = timeZoneOffset(timeZone, dtStart).replace("0", "");
	const tzNum = parseInt(tzRaw, 10);
	const hourVal = (Number.isFinite(tzNum) ? tzNum : 0) * -1 - 1;

    // End of the final day in LOCAL time
	const until = new Date(finalDay);
    until.setHours(23, 59, 59, 0);

	return {
		uid: generateUID(),
		name,
		location,
		tzid: timeZone,
		dtStart,
		dtEnd,
		until,
		byday: recurringDays,
	};
}

function serializeCalendar(events: {
    uid: string;
    name: string;
    location: string;
    tzid: string;
    dtStart: Date;
    dtEnd: Date;
    until: Date;
    byday: string[];
}[]) {
	const lines = [
		"BEGIN:VCALENDAR",
		"VERSION:2.0",
		"PRODID:-//Converted Calendar//EN",
	];

	for (const e of events) {
		lines.push(`UID:${e.uid}`);
		lines.push("BEGIN:VEVENT");
		lines.push(`SUMMARY:${escapeICalText(e.name)}`);
		lines.push(`LOCATION:${escapeICalText(e.location ?? "")}`);
		lines.push(`DTSTART;TZID=${e.tzid}:${formatDateTimeLocal(e.dtStart)}`);
		lines.push(`DTEND;TZID=${e.tzid}:${formatDateTimeLocal(e.dtEnd)}`);

		lines.push(
			`RRULE:FREQ=WEEKLY;WKST=SU;UNTIL=${formatDateTimeLocal(
				e.until
			)};BYDAY=${e.byday.join(",")}`
		);
		lines.push("END:VEVENT");
	}

	lines.push("END:VCALENDAR");
	return lines.join("\r\n");
}

function escapeICalText(s: string) {
	return String(s ?? "")
		.replace(/\\/g, "\\\\")
		.replace(/;/g, "\\;")
		.replace(/,/g, "\\,")
		.replace(/\n/g, "\\n");
}

function downloadCalendar(icsText: string, fileName: string) {
	const blob = new Blob([icsText], { type: "text/calendar;charset=utf-8" });
	const url = URL.createObjectURL(blob);

	console.log("abc");

	const hidden_link = document.createElement("a");
	hidden_link.setAttribute("download", fileName);
	hidden_link.setAttribute("href", url);
	hidden_link.click();

	setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Equivalent export to document
export default convert;
