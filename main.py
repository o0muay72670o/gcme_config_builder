"""Local CLI for preparing and applying reviewed Config worksheet mappings."""
import argparse
import json
from pathlib import Path
from config_builder.workflow import prepare_config, apply_config


def main(argv=None):
    parser=argparse.ArgumentParser(description='Build engineering Excel Config worksheets with a local Codex chat handoff. No API key.')
    commands=parser.add_subparsers(dest='command',required=True)
    prepare=commands.add_parser('prepare',help='Extract a form for reviewed field mapping')
    prepare.add_argument('source',type=Path)
    prepare.add_argument('--sheets',nargs='+',help='Exact form and supporting Condition sheet names')
    prepare.add_argument('--jobs',type=Path,default=Path('jobs'))
    apply=commands.add_parser('apply',help='Apply config-plan.json to a new workbook copy')
    apply.add_argument('request',type=Path)
    apply.add_argument('--output',type=Path,default=Path('output'))
    args=parser.parse_args(argv)
    try:
        result=prepare_config(args.source,args.jobs,args.sheets) if args.command=='prepare' else apply_config(args.request,args.output)
        print(json.dumps(result,ensure_ascii=True,indent=2))
        return 0
    except (ValueError,OSError,KeyError,TypeError) as exc:
        print(f'Error: {exc}')
        return 1


if __name__=='__main__':
    raise SystemExit(main())
