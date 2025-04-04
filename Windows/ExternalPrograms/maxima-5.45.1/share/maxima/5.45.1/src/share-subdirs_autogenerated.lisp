;; -*- Mode: lisp -*-
;;
;; Setup the share-subdirs-list.  This is a list of each subdirectory
;; (basically) of share that contains code that can be run by maxima.
;;
;; A default version is checked in, but configure will update this
;; with the correct values.  The default checked in version is useful
;; if you're doing a lisp-only build without configure.  But configure
;; is the preferred way to build.
;;
;; The file isn't directly created as share-subdirs.lisp but with an
;; immediate step in order to make the autotools not delete the file
;; we want to ship with the repo.
(in-package :maxima)

(defun share-subdirs-list ()
  '#.(remove-if-not
      #'stringp
      ;; DO NOT EDIT THIS LIST.  It is automatically
      ;; generated by configure.
      '(
"affine" "algebra" "algebra/charsets" "algebra/solver" "amatrix" "bernstein" "calculus" "cobyla" "cobyla/ex" "cobyla/lisp" "colnew" "colnew/lisp" "combinatorics" "contrib" "contrib/Eulix" "contrib/Grobner" "contrib/Zeilberger" "contrib/alt-display" "contrib/altsimp" "contrib/binsplit" "contrib/bitwise" "contrib/boolsimp" "contrib/coma" "contrib/diffequations" "contrib/diffequations/tests" "contrib/elliptic_curves" "contrib/elliptic_curves/figures" "contrib/format" "contrib/fresnel" "contrib/gentran" "contrib/gentran/man" "contrib/gentran/test" "contrib/gf" "contrib/integration" "contrib/levin" "contrib/lurkmathml" "contrib/maxima-odesolve" "contrib/maximaMathML" "contrib/mcclim" "contrib/namespaces" "contrib/noninteractive" "contrib/odes" "contrib/operatingsystem" "contrib/prim" "contrib/rand" "contrib/rkf45" "contrib/sarag" "contrib/smath" "contrib/state" "contrib/symplectic_ode" "contrib/trigtools" "contrib/unicodedata" "contrib/unit" "contrib/vector3d" "descriptive" "diff_form" "diff_form/tests" "diffequations" "distrib" "draw" "dynamics" "ezunits" "fftpack5" "fftpack5/lisp" "finance" "fourier_elim" "fractals" "graphs" "hompack" "hompack/lisp" "hypergeometric" "integequations" "integer_sequence" "integration" "lapack" "lapack/blas" "lapack/lapack" "lbfgs" "linearalgebra" "logic" "lsquares" "macro" "matrix" "minpack" "minpack/lisp" "misc" "mnewton" "multiadditive" "numeric" "numericalio" "odepack" "odepack/src" "orthopoly" "pdiff" "physics" "pytranslate" "simplex" "simplex/Tests" "simplification" "solve_rat_ineq" "solve_rec" "sound" "stats" "stringproc" "sym" "tensor" "to_poly_solve" "trigonometry" "utils" "vector" "z_transform" 
	)))
