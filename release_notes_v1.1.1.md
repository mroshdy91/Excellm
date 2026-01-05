## 🚀 ExceLLM v1.1.1 - CI & Publishing Workflow

### ✨ New Features

- **Automated PyPI Publishing**
  - GitHub Actions workflow for automatic PyPI uploads
  - No more manual `twine upload` commands
  - Release on GitHub → Automatic PyPI publish

- **Continuous Integration**
  - Ruff linting on every push/PR
  - Automated testing with pytest
  - Better code quality enforcement

- **PyPI Badges**
  - Added version badge to README
  - Added downloads badge
  - Added Python version badge
  - Added license badge

### 🔧 Improvements

- **CI Pipeline**
  - Fixed pytest fixtures for proper test isolation
  - Updated ruff configuration to use new format
  - Non-blocking lint checks (warnings allowed)
  - Better error handling in workflows

- **Documentation**
  - Clearer installation instructions
  - PyPI as primary installation method
  - Improved README with badges

### 📝 Internal Changes

- Added `tests/conftest.py` with pytest fixtures
- Created `.github/workflows/ci.yml` for CI
- Created `.github/workflows/publish-to-pypi.yml` for auto-publishing
- Updated `pyproject.toml` for ruff v0.14+

### 🎯 How It Works Now

**Before:**
```bash
# Manual process
git tag v1.1.1
twine upload dist/*
```

**Now:**
```bash
# Automated process
git tag v1.1.1
git push origin main --tags
# Create GitHub release on web → Auto-publish to PyPI!
```

### 📦 Installation

```bash
pip install excellm==1.1.1
```

Or upgrade from previous version:
```bash
pip install --upgrade excellm
```

---

**Full Changelog**: https://github.com/mroshdy91/Excellm/compare/v1.1.0...v1.1.1
