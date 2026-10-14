"""
One-click pipeline: Step 1 (discrepancy check) followed automatically by Step 2 (investor-wise sheets).
Uses exactly the same two engines as the separate tabs — nothing is re-implemented here.
"""
import os

import discrepancy_engine
import investor_engine
from reporting import UserError, log, progress, progress_range

STEP1_FILE = 'Discrepancy_Result.xlsx'
STEP2_FILE = 'Investor_Wise_Final.xlsx'


def run_both(cams_path, source_path, out_folder):
    """cams_path   : original CAMS commission Excel (Step 1 input)
       source_path : large transaction Excel for Step 2; blank/None = use the same file as cams_path
       out_folder  : both result files are written here"""
    if not cams_path or not os.path.isfile(cams_path):
        raise UserError(f'Input file not found: {cams_path}')
    source_path = source_path or cams_path
    if not os.path.isfile(source_path):
        raise UserError(f'Source transaction file not found: {source_path}')
    if not out_folder:
        raise UserError('Please choose an output folder.')
    try:
        os.makedirs(out_folder, exist_ok=True)
    except OSError as e:
        raise UserError(f'Cannot create the output folder "{out_folder}": {e}')
    step1_out = os.path.join(out_folder, STEP1_FILE)
    step2_out = os.path.join(out_folder, STEP2_FILE)

    log('══ PART 1 of 2 — Discrepancy Checker ══')
    with progress_range(0, 50):
        r1 = discrepancy_engine.run(cams_path, step1_out)
    if not r1['report_rows']:
        raise UserError(f'Part 1 finished ({STEP1_FILE} saved) but no temporary discrepancies were found, '
                        'so there is nothing for Part 2 to process.')

    log('\n══ PART 2 of 2 — Investor-wise Automation ══')
    with progress_range(50, 100):
        r2 = investor_engine.run(step1_out, source_path, step2_out)
    progress(100, 'Done')
    return {**r1, 'groups': r2['groups'], 'matched': r2['matched'], 'transactions': r2['transactions'],
            'files': r2['files'], 'step1_output': step1_out, 'output': out_folder}
