.PHONY: install dev build test test-corpus lint lint-fix clean publish

install:
	npm install

dev:
	npm run dev

build:
	npm run build

test:
	npm test

test-corpus:
	npm run test:corpus

lint:
	npm run lint

lint-fix:
	npm run lint:fix

clean:
	rm -rf dist

publish: clean build
	npm publish --access public
