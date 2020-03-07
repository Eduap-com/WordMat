;;; The Maxima test suite is defined to be the sum of all the tests
;;; in the files listed in the Maxima option variable $testsuite_files.
;;; Each entry can either be a plain file name (minus extension), e.g.,
;;; "testfile.mac", or a Maxima list containing a file name followed by
;;; the numbers of the test problems that are expected to fail, e.g.
;;; ((mlist simp) "testfile.mac" 7 9 13).

(setf $testsuite_files
      '((mlist simp)
        "rtest_rules"
        "rtestnset" 
        ((mlist) "rtest1" 180)
        "rtest1a"
        ((mlist) "rtest2" 86 95)
	"rtest4"
        "rtest5"
        "rtest6" "rtest6a" "rtest6b" "rtest7"
        "rtest9" 
        "rtest9a"
        ((mlist simp) "rtest10" 24 25)
        "rtest11" "rtest13" "rtest13s"
        "rtest14"
        "rtest15"
        "rtest16"
        "rtestode" "rtestode_zp"
        "rtest3" "rtest8"
        ((mlist simp) "rtest12" 76 78)
        "rexamples"
        ((mlist simp) "rtesthyp" 105 112 113 123 124 128)
        ((mlist simp) "rtest_hypgeo" 143)
        "rtestmt19937"
        "rtest_allnummod"
        "rtestconjugate"
        ((mlist simp) "rtestsum" 3 4 18 75)
        "rtest_trig"
        "rtest_zeta"
        "rtest_diff_invtrig"
        "rtest_scalarp"
        "rtest_everysome"
        "rtestint"
        "rtest_numth"
        "rtestifactor"
        ((mlist simp) "rtest_equal" 157 160)
        "rtest_abs"
        ((mlist simp) "rtest_taylor" 88 91 94 99 123 124)
        ((mlist simp) "rtest_dot")
        "rtest_mset"
        "rtest_boolean"
        "rtest_round"
        ((mlist simp) "rtest_map" 2 3 4)
        ((mlist simp) "rtest_sign" 21 25 30 40 65 72 79)
        "rtest_algebraic"
        "rtest_gamma"
        "rtest_expintegral"
        "rtest_signum"
        "rtest_lambert_w"
        "rtest_elliptic"
        "rtest_integrate"
        "rtest_integrate_special"
        ((mlist simp) "rtest_sqrt" 89)
        ((mlist simp) "rtest_carg" 40 41)
        ((mlist simp) "rtest_log")
        ((mlist simp) "rtest_power" 19 20 26 58 65)
        "rtestdefstruct"
        "rtest_limit"
        "rtest_powerseries"
        ((mlist) "rtest_laplace" 29 49 50 51 54 59 60 61 62 78 80)
        "rtest_plotoptions"
	))

