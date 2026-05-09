"""Test key rows with debug output."""
import sys, io, logging, traceback
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
logging.basicConfig(level=logging.WARNING, stream=sys.stderr)

try:
    from engine.fill_jobtrack import fill_jobtrack
    output, results_log, stats = fill_jobtrack(
        'Template_Files/JT_Without_copy.xlsx',
        'Template_Files/Stores_copy.xlsx',
        'Template_Files/PR_copy.xlsx',
        granules_file='Template_Files/Granules Recipe - February 2026.xlsx',
        megapack_file='Template_Files/MEGA PACK.xlsx'
    )
    print(f"Done. film={stats['film_filled']}, fresh1={stats['fresh1_filled']}, fresh2={stats['fresh2_filled']}, errors={stats['errors']}")
    for e in results_log:
        if e['row'] in [34, 46, 51, 54, 63, 64, 65]:
            print(f"Row {e['row']}: {e}")
except Exception as ex:
    traceback.print_exc()
