# nmake build
# nmake /a test
# nmake clean

# When using -g installed node-gyp
#GYP = node-gyp
# When using node.js-bundled node-gyp
GYP = node "C:\Program Files (x86)\nodejs\node_modules\npm\node_modules\node-gyp\bin\node-gyp.js"

build:
	$(GYP) configure
	$(GYP) build
	mkdir test\tmp

clean:
	$(GYP) clean
	del /Q test\tmp\*.*

test: build
	del /Q test\tmp\*.*
	set NODE_PATH=./lib;$(NODE_PATH)
	mocha -I lib test/init_win32ole.test
	mocha -I lib test/unicode.test
	node examples/maze_creator.js
	node examples/maze_solver.js

.PHONY: build test clean