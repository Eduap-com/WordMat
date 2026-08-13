CC=clang
LD=ld
CFLAGS=-g -Wall -Wundef -Wsign-compare -Wpointer-arith -O3 -g -Wall -fdollars-in-identifiers -arch arm64 -fno-omit-frame-pointer
ASFLAGS=-g -Wall -Wundef -Wsign-compare -Wpointer-arith -O3 -g -Wall -fdollars-in-identifiers -arch arm64 -fno-omit-frame-pointer
LINKFLAGS=-g -dynamic -twolevel_namespace -arch arm64
LDFLAGS=
__LDFLAGS__=
LIBS=-lc -ldl -lpthread -lzstd  -lm
