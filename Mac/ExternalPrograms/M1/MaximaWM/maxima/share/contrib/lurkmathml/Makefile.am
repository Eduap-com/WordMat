if ENABLE_BUILD_DOCS
   DOC_DIR = doc
endif

SUBDIRS = $(DOC_DIR) 
MAXIMA = $(abs_top_builddir)/maxima-local
TESTS =

if CMUCL
TESTS += mathml-test-cmucl.xml.sh
endif

if SBCL
TESTS += mathml-test-sbcl.xml.sh
endif

if GCL
TESTS += mathml-test-gcl.xml.sh
endif

if ACL
TESTS += mathml-test-acl.xml.sh
endif

if OPENMCL
TESTS += mathml-test-openmcl.xml.sh
endif

if CCL64
TESTS += mathml-test-ccl64.xml.sh
endif

if ECL
TESTS += mathml-test-ecl.xml.sh
endif

if ABCL
TESTS += mathml-test-abcl.xml.sh
endif

if CLISP
TESTS += mathml-test-clisp.xml.sh
endif

mathml-test-%.xml.sh: mathmltest.mac example.mac
	sed 's|@MAXIMA@|'"$(MAXIMA)"'|g; s|@LISP@|'$(*)'|g; s|@OUTPUT@|'$(@:.sh=)'|g;' mathmltest.sh > $@
	chmod 0755 $@

