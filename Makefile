# nmake build
# nmake /a test
# nmake clean

# When using -g installed node-gyp
#GYP = node-gyp
# When using node.js-bundled node-gyp
GYP = node "C:\Program Files (x86)\nodejs\node_modules\npm\node_modules\node-gyp\bin\node-gyp.js"

PSRC = src
HEADS0 = $(PSRC)/node_win32ole.h $(PSRC)/ole32core.h
HEADSA = $(HEADS0) $(PSRC)/v8variant.h $(PSRC)/client.h
SRCS0 = $(PSRC)/node_win32ole.cc $(PSRC)/win32ole_gettimeofday.cc
SRCS1 = $(PSRC)/client.cc $(PSRC)/v8variant.cc $(PSRC)/ole32core.cpp
SRCSA = $(SRCS0) $(SRCS1)
POBJ = build/Release/obj/node_win32ole
OBJS0 = $(POBJ)/node_win32ole.obj $(POBJ)/win32ole_gettimeofday.obj
OBJS1 = $(POBJ)/client.obj $(POBJ)/v8variant.obj $(POBJ)/ole32core.obj
OBJSA = $(OBJS0) $(OBJS1)
PTGT = build/Release
PCNF = build
TARGET = $(PTGT)/node_win32ole.node

$(TARGET) : $(PCNF)/config.gypi # $(OBJSA)
	$(GYP) rebuild

$(PCNF)/config.gypi : $(SRCSA) $(HEADSA)
	$(GYP) configure

$(POBJ)/node_win32ole.obj : $(PSRC)/$(*B).cc $(PSRC)/$(*B).h $(PSRC)/client.h $(PSRC)/v8variant.h
	$(GYP) rebuild

$(POBJ)/win32ole_gettimeofday.obj : $(PSRC)/$(*B).cc $(HEADS0)
	$(GYP) rebuild

$(POBJ)/client.obj : $(PSRC)/$(*B).cc $(PSRC)/$(*B).h $(HEADS0) $(PSRC)/v8variant.h
	$(GYP) rebuild

$(POBJ)/v8variant.obj : $(PSRC)/$(*B).cc $(PSRC)/$(*B).h $(HEADS0)
	$(GYP) rebuild

$(POBJ)/ole32core.obj : $(PSRC)/$(*B).cpp $(PSRC)/$(*B).h
	$(GYP) rebuild

build: # $(TARGET)
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

all: build test

.PHONY: build test clean
