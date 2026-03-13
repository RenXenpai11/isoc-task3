function getProjectSummary() {
	try {
		const db = getMainDatabase();
		const sheet = db.getSheetByName('Projects');
		if (!sheet) {
			return {
				success: false,
				error: "Projects sheet not found"
			};
		}

		const values = sheet.getRange('G2:K2').getValues();
		const row = (values && values[0]) ? values[0] : [];

		if (row.length < 4) {
			return {
				success: false,
				error: "Projects summary range does not contain the expected data"
			};
		}

		const [totalProjects, onTrackProjects, atRiskProjects, delayedProjects] = row;

		return {
			success: true,
			data: {
				total_projects: totalProjects,
				ontrack_projects: onTrackProjects,
				atrisk_projects: atRiskProjects,
				delayed_projects: delayedProjects
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
