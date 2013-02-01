# nmake build
# nmake /a test
# nmake clean

# When using -g installed node-gyp
#GYP = node-gyp
# When using node.js-bundled node-gyp
GYP = node "C:\Program Files (x86)\nodejs\node_modules\npm\node_modules\node-gyp\bin\node-gyp.js"

POBJ = build/Release/obj/node_win32ole

$(POBJ)/node_win32ole.obj : src/$(*B).cc src/$(*B).h src/ole32core.h
	$(GYP) rebuild

$(POBJ)/win32ole_gettimeofday.obj : src/$(*B).cc src/node_win32ole.h src/ole32core.h
	$(GYP) rebuild

$(POBJ)/ole32core.obj : src/$(*B).cpp src/$(*B).h
	$(GYP) rebuild

build:
	$(GYP) configure
	$(GYP) build
	if exist test\tmp del /Q /S test\tmp\*.*
	if not exist test\tmp mkdir test\tmp

clean:
	$(GYP) clean
	if exist test\tmp del /Q /S test\tmp\*.*
	if not exist test\tmp mkdir test\tmp

test: build
	if exist test\tmp del /Q /S test\tmp\*.*
	if not exist test\tmp mkdir test\tmp
	set NODE_PATH=./lib;$(NODE_PATH)
	mocha -I lib test/init_win32ole.test
	mocha -I lib test/unicode.test
	node examples/maze_creator.js
	node examples/maze_solver.js

all: $(POBJ)/node_win32ole.obj $(POBJ)/win32ole_gettimeofday.obj $(POBJ)/ole32core.obj build test
