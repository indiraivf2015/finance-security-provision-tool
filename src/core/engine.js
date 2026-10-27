// Builds a provision's engine from its definition. The engine has the same shape as the Security engine
// (load → analyse → build), so the shared screens (src/core/ui/) work for every provision.
//
// def = {
//   toolName, version, sopVersion,
//   extract: { accounts, accountsText, glParams, multipleAccountsAllowed },   see extract.js
//   opening: (period) => ({ columns, aprilNote }),                            see opening.js
//   provisionFile: { sheet, amount, noun },                                   see provision-file.js
//   stopInApril: false,                                                       stop for April (no year-end rule in the SOP)
//   compute: (state) => R,  buildWorkbook: (state, R, meta) => Promise<buffer>,
//   ui: { … texts and figures for the screens, see src/core/ui/app.js },
// }
'use strict';
const common = require('./common');
const { loadWorkbook } = require('./excel');
const { readExtract } = require('./extract');
const { readOpening } = require('./opening');
const { readProvisionFile } = require('./provision-file');
const { cleanLines, applyTags } = require('./clean');
const { proposeFY } = require('./tagger');

const { InputError } = common;

function createEngine(def) {
  async function load(files) {
    const errors = []; const out = {};
    for (const [k, label] of [['extract', 'ERP extract'], ['opening', 'Last month\'s statement'], ['rajeev', 'Provision file']]) {
      if (!files[k]) { errors.push(`${label}: no file selected.`); continue; }
      try { out[k] = await loadWorkbook(files[k], label); } catch (e) { if (e instanceof InputError) errors.push(...e.list); else throw e; }
    }
    if (errors.length) throw new InputError(errors);
    return out;
  }

  function analyse(wbs, opts) {
    opts = opts || {};
    const extract = readExtract(wbs.extract, def.extract);
    const period = extract.period;
    if (def.stopInApril && period.m === 4) {
      throw new InputError([`The ERP extract is for ${period.label}: April, the start of a new financial year. The SOP does not describe the year-end roll-over, so the tool cannot proceed without finance's rule.`]);
    }
    const opening = readOpening(wbs.opening, period, opts, def.opening(period));
    if (opening.needChoice) return { needChoice: opening.needChoice, period };
    const rajeev = readProvisionFile(wbs.rajeev, period, opts, def.provisionFile);
    if (rajeev.needChoice) return { needChoice: rajeev.needChoice, period };
    const { kept, removed } = cleanLines(extract.rows, period);
    return { period, extract, opening, rajeev, kept, removed };
  }

  function compute(state, tagChoice) {
    applyTags(state.kept, state.period, tagChoice);
    return def.compute(state);
  }

  async function build(state, tagChoice, meta) {
    const R = compute(state, tagChoice);
    const buffer = await def.buildWorkbook(state, R, meta);
    return { buffer, result: R };
  }

  return {
    SOP: { version: def.sopVersion }, TOOL_VERSION: def.version, toolName: def.toolName, ui: def.ui, InputError,
    load, analyse, compute, build, proposeFY, fyOf: common.fyOf, fyLabelWorking: common.fyLabelWorking, fyLabelHeader: common.fyLabelHeader, fmtINR: common.fmtINR,
  };
}

module.exports = { createEngine };
