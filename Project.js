function getProjectSummary() {
	try {
		const result = getCellValue('Projects', 'G2:K2');
		if (!result.success) {
			return result;
		}

		const values = result.values || [[result.value]];
		const row = values[0] || [];

		if (row.length < 4) {
			return {
				success: false,
				error: "Projects!G2:K2 does not contain the expected summary data"
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
