
build:
	node-gyp build

clean:
	rm test/tmp/*
	node-gyp clean

ifndef only
test: build db
	@rm -rf ./test/tmp && mkdir -p ./test/tmp
	@PATH="./node_modules/.bin:${PATH}" && NODE_PATH="./lib:$(NODE_PATH)" expresso -I lib test/*.test.js
else
test: build db
	@rm -rf ./test/tmp && mkdir -p ./test/tmp
	@PATH="./node_modules/.bin:${PATH}" && NODE_PATH="./lib:$(NODE_PATH)" expresso -I lib test/${only}.test.js
endif

.PHONY: build clean test