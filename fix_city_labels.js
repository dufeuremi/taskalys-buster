// Code à remplacer dans la fonction updateExcelSheetForGraphProv
// Remplacer la section "3. Chercher les cellules contenant des valeurs numériques dans la feuille"

// 3. Modifier les labels ET les valeurs
let modified = false;
let valueIndex = 0;
let labelIndex = 0;

// Extraire les noms des villes
const cityNames = communeMatches.map(match => match[1].trim());
addLog(`      🏙️ Noms de villes à injecter: [${cityNames.join(', ')}]`, 'info');

for (const cellRef in sheet) {
    if (cellRef.startsWith('A') || cellRef.startsWith('B') || cellRef.startsWith('C') || 
        cellRef.startsWith('D') || cellRef.startsWith('E') || cellRef.startsWith('F') ||
        cellRef.startsWith('G') || cellRef.startsWith('H')) {
        
        const cell = sheet[cellRef];
        
        // Modifier les labels (cellules contenant du texte)
        if (cell && cell.v !== undefined && typeof cell.v === 'string' && 
            (cell.v.toLowerCase().includes('ville') || cell.v.toLowerCase().includes('city'))) {
            if (labelIndex < cityNames.length) {
                const oldLabel = cell.v;
                cell.v = cityNames[labelIndex];
                addLog(`        🏙️ Label ${cellRef}: "${oldLabel}" → "${cell.v}"`, 'success');
                labelIndex++;
                modified = true;
            }
        }
        
        // Modifier les valeurs numériques
        else if (cell && cell.v !== undefined && typeof cell.v === 'number' && cell.v > 0 && cell.v < 1) {
            // C'est une valeur décimale (probablement une donnée de graphique)
            if (valueIndex < values.length) {
                const oldValue = cell.v;
                cell.v = values[valueIndex];
                addLog(`        📊 Valeur ${cellRef}: ${oldValue} → ${cell.v} (${(cell.v * 100).toFixed(1)}%)`, 'success');
                valueIndex++;
                modified = true;
            }
        }
    }
}

if (modified) {
    addLog(`      ✅ ${valueIndex} valeurs et ${labelIndex} labels modifiés dans la feuille Excel`, 'success');
} else {
    addLog(`      ⚠️ Aucune valeur ou label modifié dans la feuille Excel`, 'warning');
}
