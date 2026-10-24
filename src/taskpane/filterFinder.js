/* global Excel */

function isActiveCriterion(criterion) {
  if (!criterion || criterion.filterOn == null) return false;

  const filterType = String(criterion.filterOn).toLowerCase();
  if (filterType === "values") {
    return Array.isArray(criterion.values) && criterion.values.length > 0;
  }

  return [
    criterion.criterion1,
    criterion.criterion2,
    criterion.dynamicCriteria,
    criterion.color,
    criterion.icon,
  ].some((value) => value != null && value !== "" && value !== "Unknown");
}

function describeCriterion(criterion) {
  const values = Array.isArray(criterion.values) ? criterion.values.slice(0, 5) : criterion.values;
  return JSON.stringify({
    filterOn: criterion.filterOn,
    criterion1: criterion.criterion1,
    criterion2: criterion.criterion2,
    dynamicCriteria: criterion.dynamicCriteria,
    values,
  });
}

export async function findLastFilteredColumn() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const candidates = [];

    const worksheetFilter = sheet.autoFilter;
    worksheetFilter.load("enabled,criteria");
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    if (worksheetFilter.enabled) {
      const filterRange = worksheetFilter.getRange();
      filterRange.load("columnIndex");
      await context.sync();
      worksheetFilter.criteria.forEach((criterion, index) => {
        if (isActiveCriterion(criterion)) {
          candidates.push({
            columnIndex: filterRange.columnIndex + index,
            cell: filterRange.getCell(0, index),
            source: "worksheet filter",
            details: describeCriterion(criterion),
          });
        }
      });
    }

    for (const table of tables.items) {
      table.autoFilter.load("enabled,criteria");
    }
    await context.sync();

    for (const table of tables.items) {
      const tableFilter = table.autoFilter;
      if (tableFilter.enabled) {
        const header = table.getHeaderRowRange();
        header.load("columnIndex");
        await context.sync();
        tableFilter.criteria.forEach((criterion, index) => {
          if (isActiveCriterion(criterion)) {
            candidates.push({
              columnIndex: header.columnIndex + index,
              cell: header.getCell(0, index),
              source: `table ${table.name}`,
              details: describeCriterion(criterion),
            });
          }
        });
      }
    }

    if (candidates.length === 0) return null;

    const rightmost = candidates.reduce((last, current) =>
      current.columnIndex > last.columnIndex ? current : last
    );
    candidates.forEach((candidate) => candidate.cell.load("address"));
    rightmost.cell.select();
    await context.sync();

    return {
      address: rightmost.cell.address,
      detected: candidates.map(
        (candidate) => `${candidate.cell.address} [${candidate.source}; ${candidate.details}]`
      ),
    };
  });
}
