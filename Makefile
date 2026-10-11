.PHONY: ci test syntax image

ci:
	bash scripts/ci.sh

test:
	bash test/run.sh

syntax:
	bash scripts/syntax.sh

image:
	bash scripts/image-smoke.sh
