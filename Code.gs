// ── Configuration ──
// Everyone who should be notified of new submissions. Add/remove addresses freely.
const NOTIFY_EMAILS = [
	"you@example.com",
	// "pastor@example.com",
	// "careteam@example.com",
];

/**
 * Runs on every Google Form submission (spreadsheet-bound trigger).
 * Emails the team the full submission, and — if the person left an
 * email address — sends them a short confirmation.
 *
 * Works for any of our forms (Counselling, Prayer, …) because it reads
 * whatever questions the form has rather than fixed field names.
 *
 * To install the trigger (once per form's linked spreadsheet):
 *   1. Open the form's linked Sheet → Extensions → Apps Script
 *   2. Paste this file, set NOTIFY_EMAIL above
 *   3. Triggers (clock icon) → + Add Trigger → onFormSubmit
 *      Event source: "From spreadsheet"   Event type: "On form submit"
 *   4. Save and authorize when prompted
 */
function onFormSubmit(e) {
	const answers = e.namedValues; // { "Question title": ["answer"], … }

	// Which form this came from — used in the notification subject line.
	let formName = "Form";
	try {
		formName = e.range.getSheet().getName();
	} catch (err) {
		// running outside a real trigger (e.g. the test below)
	}

	// Turn every answer into a "Question: answer" line, skipping the
	// timestamp column and any blank answers.
	const lines = [];
	for (const question in answers) {
		if (question.toLowerCase() === "timestamp") continue;
		const value = (answers[question] || []).join(", ").trim();
		if (value) lines.push(`${question}: ${value}`);
	}

	// Pull out the submitter's name + email for personalisation.
	const name = firstAnswerMatching(answers, /name/i) || "there";
	const email = findEmail(answers);
	const timestamp = (answers["Timestamp"] || [new Date()])[0];

	const emailSubject = `New ${formName} submission`;
	const emailBody = [
		`New submission from the ${formName} form:`,
		``,
		...lines,
		``,
		`Submitted: ${timestamp}`,
		`Sheet:     ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
	].join("\n");

	// ── Notify the team (comma-separated recipients) ──
	GmailApp.sendEmail(NOTIFY_EMAILS.join(","), emailSubject, emailBody);

	// ── Confirmation to the submitter (only if they gave a real email) ──
	if (email) {
		GmailApp.sendEmail(
			email,
			"We received your request",
			`Hi ${name},\n\n` +
				`Thank you for reaching out to WordFeast Gospel Network. ` +
				`We've received your request and someone from our team will be in touch soon.\n\n` +
				`— WordFeast Gospel Network`,
		);
	}
}

// Return the first answer whose *question title* matches the pattern.
function firstAnswerMatching(answers, pattern) {
	for (const question in answers) {
		if (pattern.test(question)) {
			const value = (answers[question] || [])[0];
			if (value && value.trim()) return value.trim();
		}
	}
	return "";
}

// Scan every answer for something that looks like an email address.
// (Our "Phone or email" field may hold a phone number instead.)
function findEmail(answers) {
	const emailRe = /[^\s@]+@[^\s@]+\.[^\s@]+/;
	for (const question in answers) {
		for (const value of answers[question] || []) {
			const match = String(value).match(emailRe);
			if (match) return match[0];
		}
	}
	return "";
}

// ── Optional: test the script without a real submission ──
function testOnFormSubmit() {
	onFormSubmit({
		range: { getSheet: () => ({ getName: () => "Counselling" }) },
		namedValues: {
			Timestamp: [new Date().toISOString()],
			Name: ["Test User"],
			"Phone or email": ["test@example.com"],
			"Preferred contact": ["Phone call"],
			"What would you like to talk about?": ["This is a test submission."],
		},
	});
}
