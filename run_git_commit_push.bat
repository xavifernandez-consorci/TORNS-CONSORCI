@echo off
cd /d C:\Users\Paco\Documents\GitHub\TORNS-CONSORCI
git add Documentation/MODEL_DADES.md
git commit -m "Defineix el model de dades de la versio 3" > git_commit_push_stdout.txt 2>&1
git push origin main >> git_commit_push_stdout.txt 2>&1
echo DONE > git_commit_push_done.txt