"""
CAMS Commission Automation System — entry point.

  python main.py                          -> opens the GUI
  python main.py step1 IN.xlsx OUT.xlsx   -> Step 1 from the command line
  python main.py step2 REPORT.xlsx SOURCE.xlsx OUT.xlsx
  python main.py update IN_FOLDER_OR_FILE OUT_FOLDER     -> add Rate Change + Rate Summary to existing investor sheets
  python main.py both IN.xlsx OUT_FOLDER [SOURCE.xlsx]   -> Step 1 + Step 2 in one go
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _cli(argv):
    import reporting
    reporting.set_reporter()
    try:
        if argv[0] == 'step1' and len(argv) == 3:
            import discrepancy_engine
            discrepancy_engine.run(argv[1], argv[2])
        elif argv[0] == 'step2' and len(argv) == 4:
            import investor_engine
            investor_engine.run(argv[1], argv[2], argv[3])
        elif argv[0] == 'update' and len(argv) == 3:
            import sheet_updater
            sheet_updater.run_update(argv[1], argv[2])
        elif argv[0] == 'both' and len(argv) in (3, 4):
            import pipeline
            pipeline.run_both(argv[1], argv[3] if len(argv) == 4 else '', argv[2])
        else:
            print(__doc__)
            return 2
    except reporting.UserError as e:
        print(f'\nERROR: {e}')
        return 1
    return 0


def main():
    if len(sys.argv) > 1:
        sys.exit(_cli(sys.argv[1:]))
    import gui
    gui.launch()


if __name__ == '__main__':
    main()
