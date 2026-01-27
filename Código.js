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
        const rowEmail = paramsData[i][2] ? paramsData[i][2].toString().toLowerCase().trim() : '';
        if (rowEmail === userEmail.toLowerCase().trim()) {
            userProfile = {
                nombre: paramsData[i][0],
                correo: paramsData[i][2],
                departamento: paramsData[i][4], // Column E
                perfil: parseInt(paramsData[i][5]) || 3, // Column F
                deptUsers: paramsData.filter(r => r[4] === paramsData[i][4]).map(r => r[0]) // Same department users
            };
            break;
        }
    }
    if (!userProfile) throw new Error('Usuario no autorizado: ' + userEmail);
    // 2. Base Data
    const baseSheet = ss.getSheetByName('Base');
    if (!baseSheet) throw new Error('No se encontró la pestaña "Base"');
    const baseData = baseSheet.getDataRange().getValues();
    if (baseData.length <= 1) return { user: userProfile, baseData: [], goals: [], milestones: [], headers: [] };
    const headers = baseData[0];
    const filterIdx = 92; // Column CO
    const rows = baseData.slice(1).filter(row => {
        const val = String(row[filterIdx] || "").trim().toUpperCase();
        return val !== "RENUNCIA" && val !== "PLAN B";
    });
    // Filter by Profile
    let filteredBase = rows;
    // Filter logic for initial load: 
    // Profile 1 & 2: Full Access (Profile 2 will use frontend toggle to filter departmental view)
    if (userProfile.perfil == 1 || userProfile.perfil == 2) {
        filteredBase = rows;
    }
    // Profile 3: Individual Access
    else if (userProfile.perfil == 3) {
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
    const filtered = data.filter(row => row.some(cell => cell !== "" && cell !== null && cell !== undefined));
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
    let obsSheet = ss.getSheetByName('Observaciones Responsables');
    if (!obsSheet) {
        obsSheet = ss.insertSheet('Observaciones Responsables');
        obsSheet.appendRow(['Iniciativa', 'Código', 'Nombre', 'Observación', 'Fecha']);
    }
    const date = new Date();
    obsSheet.appendRow([
        payload.iniciativa || 'Unknown',
        payload.codigo || 'N/A',
        senderName,
        payload.observacion || '',
        date
    ]);
    return { success: true };
}
