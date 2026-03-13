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

function getProjectList() {
    try {
        const spreadsheet = getMainDatabase();
        // Try both singular and plural sheet names to avoid name mismatches
        const sheet = spreadsheet.getSheetByName('Project') || spreadsheet.getSheetByName('Projects');

        if (!sheet) {
            return {
                success: false,
                error: "Sheet 'Project' (or 'Projects') not found"
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

function doGet() {
    const payload = getProjectSummary(); 
    return ContentService
        .createTextOutput(JSON.stringify(payload))
        .setMimeType(ContentService.MimeType.JSON);
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
