function getProjectSummary() {
	try {
		const db = getMainDatabase();
		const sheet = db.getSheetByName('Projects');
		if (!sheet) {
			return { success: false, error: "Projects sheet not found" };
		}

		const data = sheet.getDataRange().getValues();
		if (!data || data.length < 2) {
			return { success: false, error: "Projects sheet has no data rows" };
		}

		const statusCol = 2; // column C (0-based index)

		let total = 0;
		let onTrack = 0;
		let atRisk = 0;
		let delayed = 0;
		.setMimeType(ContentService.MimeType.JSON);
}

function testGetProjectSummary() {
	const summary = getProjectSummary();
	console.log(summary);
	return summary;
function getProjectList() {
	try {
		const spreadsheet = getMainDatabase();
		const sheet = spreadsheet.getSheetByName('Project');

		if (!sheet) {
			return {
				success: false,
				error: "Sheet 'Project' not found"
			};
		}

		const values = sheet.getRange('G2:K2').getValues()[0];
		const progressNumber = Number(values[3]);

		return {
			success: true,
			data: {
				project_name: values[0],
				status: values[1],
				project_leader: values[2],
				progress: Number.isNaN(progressNumber) ? values[3] : progressNumber
			}
		};
	} catch (error) {
		return {
			success: false,
			error: error.message
		};
	}
}

function testGetProjectSummary() {
	const summary = getProjectSummary();
	console.log(summary);
	return summary;
}

function testGetProjectList() {
	const payload = getProjectList();
	console.log(payload);
	return payload;
}


