# ============================================================================
# Office Janitor - Development Makefile
# ============================================================================
# Cross-platform Makefile for development tasks.
# On Windows, use: make <target> (requires GNU Make, e.g., via Chocolatey)
# Or use the PowerShell scripts in scripts/ directly.
#
# Usage:
#   make help          - Show all available targets
#   make install       - Install dependencies
#   make test          - Run all tests
#   make lint          - Run all linters
#   make build         - Build standalone executable
# ============================================================================

# Project configuration
PROJECT_NAME := office-janitor
PYTHON := python
PIP := $(PYTHON) -m pip
PYTEST := $(PYTHON) -m pytest
RUFF := $(PYTHON) -m ruff
MYPY := $(PYTHON) -m mypy
BLACK := $(PYTHON) -m black
PYINSTALLER := $(PYTHON) -m PyInstaller

# Directories
SRC_DIR := src
TEST_DIR := tests
BUILD_DIR := build
DIST_DIR := dist
DOCS_DIR := docs
REPORTS_DIR := reports
VENV_DIR := .venv

# File patterns
PYTHON_FILES := $(SRC_DIR) $(TEST_DIR) oj_entry.py
SOURCE_FILES := $(SRC_DIR)/office_janitor

# Test configuration
PYTEST_ARGS := -v --tb=short
PYTEST_COVERAGE := --cov=$(SOURCE_FILES) --cov-report=html --cov-report=term-missing

# Colors for output (works in most terminals)
BOLD := $(shell tput bold 2>/dev/null || echo "")
RESET := $(shell tput sgr0 2>/dev/null || echo "")
GREEN := $(shell tput setaf 2 2>/dev/null || echo "")
YELLOW := $(shell tput setaf 3 2>/dev/null || echo "")
BLUE := $(shell tput setaf 4 2>/dev/null || echo "")
RED := $(shell tput setaf 1 2>/dev/null || echo "")

# ============================================================================
# Default target
# ============================================================================
.DEFAULT_GOAL := help

# ============================================================================
# Help
# ============================================================================
.PHONY: help
help: ## Show this help message
	@echo "$(BOLD)Office Janitor - Development Commands$(RESET)"
	@echo ""
	@echo "$(BOLD)Usage:$(RESET) make $(BLUE)<target>$(RESET)"
	@echo ""
	@echo "$(BOLD)Available targets:$(RESET)"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  $(GREEN)%-20s$(RESET) %s\n", $$1, $$2}'
	@echo ""
	@echo "$(BOLD)Examples:$(RESET)"
	@echo "  make install        # Install all dependencies"
	@echo "  make test           # Run test suite"
	@echo "  make lint           # Run all linters"
	@echo "  make build          # Build standalone executable"
	@echo "  make all            # Run full CI pipeline"

# ============================================================================
# Installation & Setup
# ============================================================================
.PHONY: install install-dev install-all venv clean-venv

venv: ## Create virtual environment
	@echo "$(BLUE)Creating virtual environment...$(RESET)"
	$(PYTHON) -m venv $(VENV_DIR)
	@echo "$(GREEN)Virtual environment created at $(VENV_DIR)$(RESET)"
	@echo "Activate with: . $(VENV_DIR)/Scripts/activate (Windows) or source $(VENV_DIR)/bin/activate (Unix)"

install: ## Install production dependencies
	@echo "$(BLUE)Installing production dependencies...$(RESET)"
	$(PIP) install --upgrade pip
	$(PIP) install -e .
	@echo "$(GREEN)Production dependencies installed$(RESET)"

install-dev: ## Install development dependencies
	@echo "$(BLUE)Installing development dependencies...$(RESET)"
	$(PIP) install --upgrade pip
	$(PIP) install -e ".[dev]"
	@echo "$(GREEN)Development dependencies installed$(RESET)"

install-all: install-dev ## Install all dependencies (alias for install-dev)

clean-venv: ## Remove virtual environment
	@echo "$(YELLOW)Removing virtual environment...$(RESET)"
ifeq ($(OS),Windows_NT)
	-rmdir /s /q $(VENV_DIR) 2>nul
else
	rm -rf $(VENV_DIR)
endif
	@echo "$(GREEN)Virtual environment removed$(RESET)"

# ============================================================================
# Code Quality - Linting
# ============================================================================
.PHONY: lint lint-ruff lint-mypy lint-black check-format

lint: lint-ruff lint-mypy check-format ## Run all linters (ruff, mypy, black)
	@echo "$(GREEN)All linting checks passed!$(RESET)"

lint-ruff: ## Run ruff linter
	@echo "$(BLUE)Running ruff...$(RESET)"
	$(RUFF) check $(PYTHON_FILES)
	@echo "$(GREEN)Ruff: OK$(RESET)"

lint-ruff-fix: ## Run ruff with auto-fix
	@echo "$(BLUE)Running ruff with auto-fix...$(RESET)"
	$(RUFF) check $(PYTHON_FILES) --fix
	@echo "$(GREEN)Ruff fixes applied$(RESET)"

lint-mypy: ## Run mypy type checker
	@echo "$(BLUE)Running mypy...$(RESET)"
	$(MYPY) $(SRC_DIR)
	@echo "$(GREEN)Mypy: OK$(RESET)"

lint-mypy-strict: ## Run mypy with strict mode
	@echo "$(BLUE)Running mypy (strict)...$(RESET)"
	$(MYPY) $(SRC_DIR) --strict
	@echo "$(GREEN)Mypy strict: OK$(RESET)"

check-format: ## Check code formatting with black (no changes)
	@echo "$(BLUE)Checking formatting with black...$(RESET)"
	$(BLACK) --check --diff $(PYTHON_FILES)
	@echo "$(GREEN)Black: OK$(RESET)"

# ============================================================================
# Code Quality - Formatting
# ============================================================================
.PHONY: format format-black format-ruff format-all

format: format-black format-ruff ## Format code (black + ruff)
	@echo "$(GREEN)Code formatted!$(RESET)"

format-black: ## Format code with black
	@echo "$(BLUE)Formatting with black...$(RESET)"
	$(BLACK) $(PYTHON_FILES)
	@echo "$(GREEN)Black formatting applied$(RESET)"

format-ruff: ## Sort imports with ruff
	@echo "$(BLUE)Sorting imports with ruff...$(RESET)"
	$(RUFF) check $(PYTHON_FILES) --fix --select I
	@echo "$(GREEN)Import sorting applied$(RESET)"

format-all: format ## Alias for format

# ============================================================================
# Testing
# ============================================================================
.PHONY: test test-fast test-verbose test-coverage test-unit test-integration test-watch test-failed

test: ## Run all tests
	@echo "$(BLUE)Running tests...$(RESET)"
	$(PYTEST) $(TEST_DIR) $(PYTEST_ARGS)
	@echo "$(GREEN)All tests passed!$(RESET)"

test-fast: ## Run tests without verbose output (faster)
	@echo "$(BLUE)Running tests (fast mode)...$(RESET)"
	$(PYTEST) $(TEST_DIR) -q --tb=no
	@echo "$(GREEN)Tests passed!$(RESET)"

test-verbose: ## Run tests with extra verbose output
	@echo "$(BLUE)Running tests (verbose)...$(RESET)"
	$(PYTEST) $(TEST_DIR) -vv --tb=long

test-coverage: ## Run tests with coverage report
	@echo "$(BLUE)Running tests with coverage...$(RESET)"
	$(PYTEST) $(TEST_DIR) $(PYTEST_COVERAGE)
	@echo "$(GREEN)Coverage report generated in htmlcov/$(RESET)"

test-unit: ## Run unit tests only (exclude integration)
	@echo "$(BLUE)Running unit tests...$(RESET)"
	$(PYTEST) $(TEST_DIR) -v -m "not integration"

test-integration: ## Run integration tests only
	@echo "$(BLUE)Running integration tests...$(RESET)"
	$(PYTEST) $(TEST_DIR) -v -m "integration"

test-failed: ## Re-run only failed tests
	@echo "$(BLUE)Re-running failed tests...$(RESET)"
	$(PYTEST) $(TEST_DIR) --lf -v

test-watch: ## Run tests in watch mode (requires pytest-watch)
	@echo "$(BLUE)Starting test watch mode...$(RESET)"
	$(PYTHON) -m pytest_watch -- $(TEST_DIR) -v

test-parallel: ## Run tests in parallel (requires pytest-xdist)
	@echo "$(BLUE)Running tests in parallel...$(RESET)"
	$(PYTEST) $(TEST_DIR) -n auto -v

# Test specific modules
test-detect: ## Run detection tests
	$(PYTEST) $(TEST_DIR)/test_detect.py -v

test-scrub: ## Run scrub tests
	$(PYTEST) $(TEST_DIR)/test_scrub.py -v

test-registry: ## Run registry tests
	$(PYTEST) $(TEST_DIR)/test_registry_tools.py -v

test-msi: ## Run MSI component tests
	$(PYTEST) $(TEST_DIR)/test_msi_components.py -v

test-appx: ## Run AppX uninstall tests
	$(PYTEST) $(TEST_DIR)/test_appx_uninstall.py -v

# ============================================================================
# Building
# ============================================================================
.PHONY: build build-exe build-dist build-wheel build-sdist clean-build

build: build-exe ## Build standalone executable (PyInstaller)

build-exe: ## Build standalone executable with PyInstaller
	@echo "$(BLUE)Building standalone executable...$(RESET)"
	$(PYINSTALLER) office-janitor.spec --noconfirm
	@echo "$(GREEN)Executable built: dist/office-janitor.exe$(RESET)"

build-dist: build-wheel build-sdist ## Build distribution packages (wheel + sdist)
	@echo "$(GREEN)Distribution packages built in dist/$(RESET)"

build-wheel: ## Build wheel package
	@echo "$(BLUE)Building wheel...$(RESET)"
	$(PYTHON) -m build --wheel
	@echo "$(GREEN)Wheel built$(RESET)"

build-sdist: ## Build source distribution
	@echo "$(BLUE)Building source distribution...$(RESET)"
	$(PYTHON) -m build --sdist
	@echo "$(GREEN)Source distribution built$(RESET)"

clean-build: ## Clean build artifacts
	@echo "$(YELLOW)Cleaning build artifacts...$(RESET)"
ifeq ($(OS),Windows_NT)
	-rmdir /s /q $(BUILD_DIR) 2>nul
	-rmdir /s /q $(DIST_DIR) 2>nul
	-rmdir /s /q *.egg-info 2>nul
	-del /q *.spec 2>nul
else
	rm -rf $(BUILD_DIR) $(DIST_DIR) *.egg-info
endif
	@echo "$(GREEN)Build artifacts cleaned$(RESET)"

# ============================================================================
# Running
# ============================================================================
.PHONY: run run-help run-detect run-detect-json run-tui run-dry

run: ## Run office-janitor (show help)
	$(PYTHON) -m office_janitor --help

run-help: ## Show detailed CLI help
	$(PYTHON) -m office_janitor --help

run-detect: ## Run detection only (dry-run)
	$(PYTHON) -m office_janitor detect --dry-run

run-detect-json: ## Run detection with JSON output
	$(PYTHON) -m office_janitor detect --json --dry-run

run-tui: ## Run interactive TUI mode
	$(PYTHON) -m office_janitor --tui

run-dry: ## Run full scrub in dry-run mode
	$(PYTHON) -m office_janitor scrub --dry-run --all

run-version: ## Show version information
	$(PYTHON) -m office_janitor --version

# ============================================================================
# Documentation
# ============================================================================
.PHONY: docs docs-serve docs-build

docs: ## Generate documentation (if using mkdocs/sphinx)
	@echo "$(BLUE)Documentation is in $(DOCS_DIR)/$(RESET)"
	@echo "Available docs:"
	@ls -1 $(DOCS_DIR)/*.md 2>/dev/null || dir $(DOCS_DIR)\*.md

docs-api: ## Generate API documentation
	@echo "$(BLUE)Generating API docs...$(RESET)"
	$(PYTHON) -m pydoc -w $(SOURCE_FILES)
	@echo "$(GREEN)API docs generated$(RESET)"

# ============================================================================
# Reports
# ============================================================================
.PHONY: report report-lint report-test report-type report-all

report: report-all ## Generate all reports

report-lint: ## Generate lint report
	@echo "$(BLUE)Generating lint report...$(RESET)"
	-$(RUFF) check $(PYTHON_FILES) --output-format=text > $(REPORTS_DIR)/lint_report.txt 2>&1
	@echo "$(GREEN)Lint report: $(REPORTS_DIR)/lint_report.txt$(RESET)"

report-test: ## Generate test report
	@echo "$(BLUE)Generating test report...$(RESET)"
	-$(PYTEST) $(TEST_DIR) -v --tb=short > $(REPORTS_DIR)/test_report.txt 2>&1
	@echo "$(GREEN)Test report: $(REPORTS_DIR)/test_report.txt$(RESET)"

report-type: ## Generate type check report
	@echo "$(BLUE)Generating type check report...$(RESET)"
	-$(MYPY) $(SRC_DIR) > $(REPORTS_DIR)/type_check.txt 2>&1
	@echo "$(GREEN)Type check report: $(REPORTS_DIR)/type_check.txt$(RESET)"

report-all: report-lint report-test report-type ## Generate all reports
	@echo "$(GREEN)All reports generated in $(REPORTS_DIR)/$(RESET)"

# ============================================================================
# Cleaning
# ============================================================================
.PHONY: clean clean-pyc clean-test clean-cache clean-all

clean: clean-pyc clean-test clean-cache ## Clean common artifacts
	@echo "$(GREEN)Cleaned!$(RESET)"

clean-pyc: ## Remove Python bytecode files
	@echo "$(YELLOW)Removing Python bytecode...$(RESET)"
ifeq ($(OS),Windows_NT)
	-for /r . %%f in (*.pyc) do del "%%f" 2>nul
	-for /r . %%d in (__pycache__) do rmdir /s /q "%%d" 2>nul
else
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
endif

clean-test: ## Remove test artifacts
	@echo "$(YELLOW)Removing test artifacts...$(RESET)"
ifeq ($(OS),Windows_NT)
	-rmdir /s /q .pytest_cache 2>nul
	-rmdir /s /q htmlcov 2>nul
	-rmdir /s /q .coverage 2>nul
else
	rm -rf .pytest_cache htmlcov .coverage
endif

clean-cache: ## Remove cache directories
	@echo "$(YELLOW)Removing cache directories...$(RESET)"
ifeq ($(OS),Windows_NT)
	-rmdir /s /q .mypy_cache 2>nul
	-rmdir /s /q .ruff_cache 2>nul
else
	rm -rf .mypy_cache .ruff_cache
endif

clean-all: clean clean-build clean-venv ## Remove all generated files
	@echo "$(GREEN)All artifacts cleaned!$(RESET)"

# ============================================================================
# Development Workflows
# ============================================================================
.PHONY: all ci pre-commit check quick

all: lint test build ## Run full CI pipeline (lint, test, build)
	@echo "$(GREEN)$(BOLD)Full CI pipeline completed successfully!$(RESET)"

ci: lint test ## Run CI checks (lint + test, no build)
	@echo "$(GREEN)CI checks passed!$(RESET)"

pre-commit: format lint test-fast ## Pre-commit workflow (format, lint, quick test)
	@echo "$(GREEN)Pre-commit checks passed!$(RESET)"

check: lint test-fast ## Quick check (lint + fast tests)
	@echo "$(GREEN)Quick checks passed!$(RESET)"

quick: test-fast ## Alias for fast tests

# ============================================================================
# Git Helpers
# ============================================================================
.PHONY: git-status git-diff git-log

git-status: ## Show git status
	git status

git-diff: ## Show unstaged changes
	git diff

git-log: ## Show recent commits
	git log --oneline -10

# ============================================================================
# Utility Targets
# ============================================================================
.PHONY: info deps outdated upgrade

info: ## Show project information
	@echo "$(BOLD)Project: $(PROJECT_NAME)$(RESET)"
	@echo "Python: $(shell $(PYTHON) --version)"
	@echo "Pip: $(shell $(PIP) --version)"
	@echo "Working directory: $(shell pwd)"
	@echo ""
	@echo "$(BOLD)Source files:$(RESET)"
	@find $(SRC_DIR) -name "*.py" 2>/dev/null | wc -l || dir /s /b $(SRC_DIR)\*.py 2>nul | find /c /v ""
	@echo ""
	@echo "$(BOLD)Test files:$(RESET)"
	@find $(TEST_DIR) -name "*.py" 2>/dev/null | wc -l || dir /s /b $(TEST_DIR)\*.py 2>nul | find /c /v ""

deps: ## List installed dependencies
	$(PIP) list

outdated: ## Show outdated packages
	$(PIP) list --outdated

upgrade: ## Upgrade all packages
	@echo "$(BLUE)Upgrading packages...$(RESET)"
	$(PIP) list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1 | xargs -n1 $(PIP) install -U 2>/dev/null || \
		$(PIP) install --upgrade pip setuptools wheel
	@echo "$(GREEN)Packages upgraded$(RESET)"

loc: ## Count lines of code
	@echo "$(BOLD)Lines of Code:$(RESET)"
	@find $(SRC_DIR) -name "*.py" -exec cat {} + 2>/dev/null | wc -l || \
		powershell -Command "(Get-ChildItem -Path $(SRC_DIR) -Filter *.py -Recurse | Get-Content | Measure-Object -Line).Lines"

# ============================================================================
# Security
# ============================================================================
.PHONY: security security-check

security: security-check ## Run security checks

security-check: ## Check for security vulnerabilities (requires pip-audit)
	@echo "$(BLUE)Running security audit...$(RESET)"
	-$(PYTHON) -m pip_audit
	@echo "$(GREEN)Security check complete$(RESET)"

# ============================================================================
# Release
# ============================================================================
.PHONY: release-check release-patch release-minor release-major

release-check: lint test ## Pre-release checks
	@echo "$(GREEN)Release checks passed!$(RESET)"

release-patch: release-check ## Bump patch version (requires bump2version)
	@echo "$(BLUE)Bumping patch version...$(RESET)"
	bump2version patch

release-minor: release-check ## Bump minor version
	@echo "$(BLUE)Bumping minor version...$(RESET)"
	bump2version minor

release-major: release-check ## Bump major version
	@echo "$(BLUE)Bumping major version...$(RESET)"
	bump2version major

# ============================================================================
# Docker (if applicable)
# ============================================================================
.PHONY: docker-build docker-run docker-clean

docker-build: ## Build Docker image
	@echo "$(BLUE)Building Docker image...$(RESET)"
	docker build -t $(PROJECT_NAME):latest .

docker-run: ## Run in Docker container
	docker run --rm -it $(PROJECT_NAME):latest

docker-clean: ## Remove Docker artifacts
	docker rmi $(PROJECT_NAME):latest 2>/dev/null || true

# ============================================================================
# End of Makefile
# ============================================================================
