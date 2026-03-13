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

		for (let i = 1; i < data.length; i++) {
			const row = data[i];
			const statusRaw = (row[statusCol] + '').trim().toLowerCase();
			if (!statusRaw) {
				continue;
			}
			total++;
			if (statusRaw === 'on track') {
				onTrack++;
			} else if (statusRaw === 'at risk') {
				atRisk++;
			} else if (statusRaw === 'delayed') {
				delayed++;
			}
		}

		return {
			success: true,
			data: {
				total_projects: total,
				ontrack_projects: onTrack,
				atrisk_projects: atRisk,
				delayed_projects: delayed
			}
		};
	} catch (error) {
		const errorResult = {
			success: false,
			error: error.message
		};
		console.log(errorResult);
		return errorResult;
	}
}

function testGetProjectSummary() {
	const summary = getProjectSummary();
	console.log(summary);
	return summary;
}
