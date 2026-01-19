function doGet() {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('EVA - Dashboard de Eficiencia')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getInitialData() {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 1. Check user profile
    const paramsSheet = ss.getSheetByName('Parámetros');
    if (!paramsSheet) throw new Error('No se encontró la pestaña "Parámetros"');
    const paramsData = paramsSheet.getDataRange().getValues();
    let userProfile = null;
    for (let i = 1; i < paramsData.length; i++) {
        if (paramsData[i][2] === userEmail) {
            userProfile = {
                nombre: paramsData[i][0],
                correo: paramsData[i][2],
                perfil: parseInt(paramsData[i][5]) || 3
            };
            break;
        }
    }
    if (!userProfile) throw new Error('Usuario no autorizado: ' + userEmail);
    // 2. Base Data
    const baseSheet = ss.getSheetByName('Base');
    if (!baseSheet) throw new Error('No se encontró la pestaña "Base"');
    const baseData = baseSheet.getDataRange().getValues();
    // Safety check for empty base
    if (baseData.length <= 1) return { user: userProfile, baseData: [], goals: [], milestones: [], headers: [] };
    const headers = baseData[0];
    const rows = baseData.slice(1);
    // Filter by Profile
    let filteredBase = rows;
    if (userProfile.perfil == 3) {
        filteredBase = rows.filter(row => {
            return row[14] === userProfile.nombre || row[15] === userProfile.nombre || row[16] === userProfile.nombre;
        });
    }
    const goalsSheet = ss.getSheetByName('Metas VPEs');
    const goalsData = goalsSheet ? pruneEmptyRows(goalsSheet.getDataRange().getValues()) : [];
    const milestonesSheet = ss.getSheetByName('Hitos Iniciativas');
    const milestonesData = milestonesSheet ? pruneEmptyRows(milestonesSheet.getDataRange().getValues()) : [];
    return {
        user: userProfile,
        baseData: pruneEmptyRows(filteredBase),
        goals: goalsData,
        milestones: milestonesData,
        headers: headers.map(h => h ? h.toString() : "")
    };
}
function pruneEmptyRows(data) {
    if (!data || data.length === 0) return [];
    // 1. Remove rows that are entirely empty
    const filtered = data.filter(row => row.some(cell => cell !== "" && cell !== null && cell !== undefined));
    // 2. Sanitize all cells to simple primitives (String/Number/Boolean/Null)
    // This prevents serialization errors with problematic Date objects or complex types
    return filtered.map(row => row.map(cell => {
        if (cell === "" || cell === null || cell === undefined) return null;
        if (cell instanceof Date) {
            return isNaN(cell.getTime()) ? "Fecha Inválida" : cell.toISOString();
        }
        if (typeof cell === 'object') return JSON.stringify(cell);
        return cell;
    }));
}
function submitObservation(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userEmail = Session.getActiveUser().getEmail();
    // 1. Get user name from Parámetros for Column C
    let senderName = payload.nombre || 'Desconocido';
    const paramsSheet = ss.getSheetByName('Parámetros');
    if (paramsSheet) {
        const paramsData = paramsSheet.getDataRange().getValues();
        const searchEmail = userEmail.toLowerCase().trim();
        for (let i = 1; i < paramsData.length; i++) {
            const rowEmail = paramsData[i][2] ? paramsData[i][2].toString().toLowerCase().trim() : '';
            if (rowEmail === searchEmail) {
                senderName = paramsData[i][0];
                break;
            }
        }
    }
    // 2. Get/Create Observations sheet
    let obsSheet = ss.getSheetByName('Observaciones Responsables');
    if (!obsSheet) {
        obsSheet = ss.insertSheet('Observaciones Responsables');
        obsSheet.appendRow(['Iniciativa', 'Código', 'Nombre', 'Observación', 'Fecha']);
    }
    // 3. Record entry
    const date = new Date();
    // A: Iniciativa, B: Código, C: Nombre, D: Observación, E: Fecha
    obsSheet.appendRow([
        payload.iniciativa || 'Unknown',
        payload.codigo || 'N/A',
        senderName,
        payload.observacion || '',
        date
    ]);
    return { success: true };
}
