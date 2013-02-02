# nmake build
# nmake /a test
# nmake clean

# When using -g installed node-gyp
#GYP = node-gyp
# When using node.js-bundled node-gyp
GYP = node "C:\Program Files (x86)\nodejs\node_modules\npm\node_modules\node-gyp\bin\node-gyp.js"

PSRC = src
HEADS = $(PSRC)/node_win32ole.h $(PSRC)/ole32core.h
POBJ = build/Release/obj/node_win32ole

$(POBJ)/node_win32ole.node : $(POBJ)/node_win32ole.obj $(POBJ)/win32ole_gettimeofday.obj $(POBJ)/statement.obj $(POBJ)/v8variant.obj $(POBJ)/ole32core.obj
	$(GYP) rebuild

$(POBJ)/node_win32ole.obj : $(PSRC)/$(*B).cc $(PSRC)/$(*B).h $(PSRC)/statement.h
	$(GYP) rebuild

$(POBJ)/win32ole_gettimeofday.obj : $(PSRC)/$(*B).cc $(HEADS)
	$(GYP) rebuild

$(POBJ)/statement.obj : $(PSRC)/$(*B).cc $(PSRC)/$(*B).h $(HEADS) $(PSRC)/v8variant.h
	$(GYP) rebuild

$(POBJ)/v8variant.obj : $(PSRC)/$(*B).cc $(PSRC)/$(*B).h $(HEADS)
	$(GYP) rebuild

$(POBJ)/ole32core.obj : $(PSRC)/$(*B).cpp $(PSRC)/$(*B).h
	$(GYP) rebuild

all: build test

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

.PHONY: build test clean
