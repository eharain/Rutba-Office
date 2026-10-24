export { Base, DESIGN_VERSION } from './base.js';
export { FIELD_TYPES, columnSql, toStored } from './types.js';
export { importRows, importSheets, inferType } from './import.js';
export { designToSql, criterionSql, joinsFor, splitNamedExpression } from './query-design.js';
export { expressionSql, fieldsNamed, FUNCTION_NAMES } from './expression.js';
export { wizardForm, labelFor, shownField, valueFromControl, GRID } from './form-design.js';
export { wizardReport, reportData, total, totalFields, withTotal } from './report-design.js';
export { TEMPLATE_IDS, templateTables, buildTemplate, buildFromSpec } from './templates.js';
export { documentData, wizardDocument, evaluate, PAPERS, BLOCKS } from './document-design.js';
export { buildSample, buildSampleFrom } from './samples.js';
// The ready-made databases' content, from the templates' own repository.
export { SAMPLE_IDS, sampleSpec, sampleAbout, checkSample } from '@rutba/templates/data-files/active';
export { originOf, recordOrigin, designPrint } from './origin.js';
export { planUpgrade, applyUpgrade } from './upgrade.js';
// Synced copies of engines' tables, and the other sources a table is read from (D-2026-10-22-13, steps 4 and 5).
export { SOURCE_KINDS, SOURCE_EVERY, DEFAULT_EVERY, ensureSync, prepareSync, pushed, refused, reconcile, replaceRows, pending, clashList, clashCount, dropClash, designDiff, markSynced, stampMax, hashRows, isSource, copyReadOnly } from './sync.js';
export { tableFromRows, tableFromRecords, flatten, valueAt, sourceFieldNames } from './sources.js';
