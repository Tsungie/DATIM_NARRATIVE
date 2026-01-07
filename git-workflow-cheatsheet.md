# Git Workflow Cheatsheet

## ✅ Initialize a New Git Repository and Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<username>/<repo>.git
git push -u origin main
```

## ✅ Create and Push a New Branch
```bash
git checkout -b feature/demo-mode
git push -u origin feature/demo-mode
```

## ✅ Check Branches
```bash
git branch -a
```

## ✅ Merge Branch into Main
```bash
git checkout main
git merge feature/demo-mode
git push origin main
```

---

# Python .gitignore Template

```
# Byte-compiled / optimized / DLL files
__pycache__/
*.py[cod]
*$py.class

# C extensions
*.so

# Distribution / packaging
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
share/python-wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST

# Virtual environments
venv/
ENV/
env/
.venv/

# PyInstaller
*.manifest
*.spec

# Installer logs
pip-log.txt
pip-delete-this-directory.txt

# Unit test / coverage reports
htmlcov/
.tox/
.nox/
.coverage
.coverage.*
.cache
nosetests.xml
coverage.xml
*.cover
*.py,cover
.hypothesis/
.pytest_cache/

# Jupyter Notebook
.ipynb_checkpoints

# pyenv
.python-version

# mypy
.mypy_cache/
.dmypy.json
dmypy.json

# VS Code
.vscode/

# macOS
.DS_Store

# Logs
*.log

# Local config
*.env
```
