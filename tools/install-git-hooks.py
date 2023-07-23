from pathlib import Path
import os
import stat

PRE_COMMIT_CONTENT = """\
#!/bin/sh

# python .git/hooks/pre-commit.py
echo ./*.vba
"""

def install_hooks():
    base_path = Path(__file__).parent / "../.git/hooks"
    pre_commit_path = (base_path / "pre-commit").resolve()

    with open(pre_commit_path, "w") as fis:
        fis.write(PRE_COMMIT_CONTENT)
        st = os.stat(pre_commit_path)
        os.chmod(pre_commit_path, st.st_mode | stat.S_IEXEC)

    print(f"Written {pre_commit_path} successfully !")

if __name__ == '__main__':
    install_hooks()
