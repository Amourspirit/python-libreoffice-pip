help:
	@echo "Run make build to compile into 'build' folder."
	@echo "Run make build_idl to compile idl"
	@echo "Run make diff to write 'changes_since_last_commit.diff into' into 'tmp' folder."
	@echo "Run make diff_sum to write 'summary_since_last_commit.diff into' into 'tmp' folder."

.PHONY: build build_idl diff diff_sum help

build:
	uv run --no-config make.py build --no-idl

build_idl:
	uv run --no-config make.py build

create_build_dir:
	mkdir -p tmp

diff: create_build_dir
	git diff > ./tmp/changes_since_last_commit.diff

diff_sum: create_build_dir
	git diff --compact-summary > ./tmp/summary_since_last_commit.diff