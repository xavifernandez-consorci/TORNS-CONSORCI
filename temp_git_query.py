from pathlib import Path
from subprocess import check_output, CalledProcessError

base = Path(__file__).resolve().parent
out = base / 'temp_git_query_output.txt'
try:
    rev = check_output(['git', 'rev-parse', 'HEAD'], cwd=base, encoding='utf-8').strip()
    log = check_output(['git', 'log', '--oneline', '-5'], cwd=base, encoding='utf-8').strip()
    status = check_output(['git', 'status', '--short'], cwd=base, encoding='utf-8').strip()
    out.write_text(f'REV:\n{rev}\n\nLOG:\n{log}\n\nSTATUS:\n{status}', encoding='utf-8')
except CalledProcessError as ex:
    out.write_text(f'ERROR: {ex}\nSTDOUT:\n{ex.output}\nSTDERR:\n{ex.stderr}', encoding='utf-8')
