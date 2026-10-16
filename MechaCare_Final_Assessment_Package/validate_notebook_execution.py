import subprocess
import sys
from pathlib import Path

nb_path = Path(r'C:\Users\Hari Vignesh S\Downloads\MechaCare_Final_Assessment_Package1\MechaCare_Final_Assessment_Package\MechaCare_EDA_Feature_Engineering.ipynb')
cmd = [sys.executable, '-m', 'jupyter', 'nbconvert', '--to', 'notebook', '--execute', '--inplace', str(nb_path)]
print('RUNNING:', ' '.join(cmd))
result = subprocess.run(cmd, capture_output=True, text=True)
print('RETURN_CODE:', result.returncode)
if result.stdout:
    print('STDOUT:\n' + result.stdout[-2000:])
if result.stderr:
    print('STDERR:\n' + result.stderr[-2000:])
if result.returncode != 0:
    raise SystemExit(result.returncode)
print('NOTEBOOK_VALIDATE_OK')
