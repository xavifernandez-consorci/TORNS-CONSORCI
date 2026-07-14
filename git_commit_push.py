import subprocess
from pathlib import Path

base = Path(__file__).resolve().parent
log = base / 'git_commit_push_log.txt'

commands = [
    ['git', 'add', 'Excel/TORNS-CONSORCI.xlsm', '.gitignore'],
    ['git', 'commit', '-m', "Afegeix el llibre host de l'aplicació"],
    ['git', 'push', 'origin', 'main'],
]
output_lines = []
for cmd in commands:
    output_lines.append(f'COMMAND: {" ".join(cmd)}')
    try:
        result = subprocess.run(cmd, cwd=base, capture_output=True, text=True, check=True)
        output_lines.append('RETURNCODE: 0')
        if result.stdout:
            output_lines.append('STDOUT:')
            output_lines.append(result.stdout.rstrip())
        if result.stderr:
            output_lines.append('STDERR:')
            output_lines.append(result.stderr.rstrip())
    except subprocess.CalledProcessError as exc:
        output_lines.append(f'RETURNCODE: {exc.returncode}')
        if exc.stdout:
            output_lines.append('STDOUT:')
            output_lines.append(exc.stdout.rstrip())
        if exc.stderr:
            output_lines.append('STDERR:')
            output_lines.append(exc.stderr.rstrip())
        break

log.write_text('\n'.join(output_lines), encoding='utf-8')
