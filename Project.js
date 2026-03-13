function getProjectSummary() {
	try {
		const spreadsheet = getMainDatabase();
		const sheet = spreadsheet.getSheetByName('Projects');

		if (!sheet) {
			return {
				success: false,
				error: "Sheet 'Projects' not found"
			};
		}

		const rows = sheet.getDataRange().getValues();
		const summary = {
			total_projects: 0,
			ontrack_projects: 0,
			atrisk_projects: 0,
			delayed_projects: 0
		};

		if (rows.length <= 1) {
			return {
				success: true,
				data: summary
			};
		}

		const headers = rows[0].map(header => String(header).trim().toLowerCase());
		const statusIndex = headers.findIndex(header =>
			header === 'status' ||
			header === 'project status' ||
			header === 'project_status' ||
			header === 'health' ||
			header === 'progress'
		);

		if (statusIndex === -1) {
			return {
				success: false,
				error: "Status column not found in 'Projects' sheet"
			};
		}

		rows.slice(1).forEach(row => {
			const hasData = row.some(cell => cell !== '' && cell !== null);
			if (!hasData) return;

			summary.total_projects += 1;

			const normalizedStatus = String(row[statusIndex] || '')
				.trim()
				.toLowerCase()
				.replace(/[\s_-]+/g, '');

			if (normalizedStatus === 'ontrack') summary.ontrack_projects += 1;
			if (normalizedStatus === 'atrisk') summary.atrisk_projects += 1;
			if (normalizedStatus === 'delayed') summary.delayed_projects += 1;
		});

		return {
			success: true,
			data: summary
		};
	} catch (error) {
		return {
			success: false,
			error: error.message
		};
	}
}

function doGet() {
	const payload = getProjectSummary(); 
	return ContentService
		.createTextOutput(JSON.stringify(payload))
		.setMimeType(ContentService.MimeType.JSON);
}

function testProjectSummary() {
	const payload = getProjectSummary();
	Logger.log(JSON.stringify(payload, null, 2));
	return payload;
}
