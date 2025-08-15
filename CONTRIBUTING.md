# Contributing Guide

Thank you for your interest in contributing! 🎉  
We welcome contributions from everyone, whether it's fixing a typo, reporting a bug, or building a new feature.

## 📌 How to Contribute

### 1. Fork & Clone
1. **Fork** the repository to your GitHub account.
2. **Clone** your fork:
   ```bash
   git clone git@github.com:didikprabowo/mbadocx.git
   cd mbadocx
   ```

### 2. Create a Branch
Always create a new branch for your work:
```bash
git checkout -b feature/your-feature-name
```
Branch naming conventions:
- `feature/...` → for new features  
- `fix/...` → for bug fixes  
- `docs/...` → for documentation changes

### 3. Make Changes
- Follow the coding style used in the project.
- Keep commits **small and descriptive**.
- Write meaningful commit messages:
  ```
  feat: add ability to export to PDF
  fix: correct null pointer in DocumentParser
  docs: update README with installation steps
  ```

### 4. Run Tests
Before pushing changes, ensure tests pass:
```bash
go test ./...
```
(or other test commands if applicable)

### 5. Push & Create a Pull Request (PR)
1. Push your branch:
   ```bash
   git push origin feature/your-feature-name
   ```
2. Open a PR from your fork to the `main` branch of the original repo.
3. Fill out the PR template (if available).

---

## 🐞 Reporting Bugs
- Use [GitHub Issues](../../issues) to report bugs.
- Include:
  - Steps to reproduce
  - Expected behavior
  - Screenshots or logs if helpful
  - Your OS, Go version, or environment details

---

## 📜 Code Style
- Use **Go formatting**:
  ```bash
  go fmt ./...
  ```
- Avoid unused imports and variables.
- Keep functions short and focused.

---

Thank you for helping improve this project! 🚀
